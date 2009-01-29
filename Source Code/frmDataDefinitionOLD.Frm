VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDataDefinition 
   AutoRedraw      =   -1  'True
   Caption         =   "Question Definition"
   ClientHeight    =   6480
   ClientLeft      =   1710
   ClientTop       =   1470
   ClientWidth     =   8550
   ClipControls    =   0   'False
   Icon            =   "frmDataDefinition.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6480
   ScaleWidth      =   8550
   Begin VB.CommandButton cmdValidationExpression 
      BackColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   7200
      Picture         =   "frmDataDefinition.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   5640
      Visible         =   0   'False
      Width           =   600
   End
   Begin MSComctlLib.ImageList ilValidationTypes 
      Left            =   6480
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageCombo icboValidationType 
      Height          =   330
      Left            =   5400
      TabIndex        =   40
      Top             =   5760
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Locked          =   -1  'True
      Text            =   "Image Combo Control"
      ImageList       =   "ilValidationTypes"
   End
   Begin VB.TextBox txtValidationGridEdits 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   6000
      TabIndex        =   38
      Text            =   "validation grid editing text box"
      Top             =   5760
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtCategoryEdit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   5760
      TabIndex        =   45
      Text            =   "categories grid editing text box"
      Top             =   5880
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog cdImport 
      Left            =   3000
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ValueData"
      Top             =   5760
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   6000
      Width           =   1215
   End
   Begin TabDlg.SSTab sstDataDefinition 
      Height          =   5880
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   10372
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Definition"
      TabPicture(0)   =   "frmDataDefinition.frx":074C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmData"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Validation"
      TabPicture(1)   =   "frmDataDefinition.frx":0768
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "gridValidation"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Values"
      TabPicture(2)   =   "frmDataDefinition.frx":0784
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(1)=   "txtAfterCategroyGridTabOrder"
      Tab(2).Control(2)=   "txtAfterCategoryTextTabOrder"
      Tab(2).Control(3)=   "txtAfterValidationTextTabOrder"
      Tab(2).Control(4)=   "gridCategories"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Required"
      TabPicture(3)   =   "frmDataDefinition.frx":07A0
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lsvRequired"
      Tab(3).ControlCount=   1
      Begin VB.Frame frmData 
         Height          =   5400
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   8000
         Begin VB.Frame fraCoding 
            Caption         =   "Dictionary"
            Height          =   855
            Left            =   4680
            TabIndex        =   66
            Top             =   1620
            Width           =   3195
            Begin VB.ComboBox cboDictionary 
               Height          =   315
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   67
               Top             =   360
               Width           =   2895
            End
         End
         Begin VB.TextBox txtDescription 
            Height          =   660
            Left            =   1800
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   3840
            Width           =   6075
         End
         Begin VB.CheckBox chkMACROOnly 
            Caption         =   "Do not send to Oracle Clinical"
            Height          =   255
            Left            =   1800
            TabIndex        =   3
            Top             =   1320
            Width           =   2775
         End
         Begin VB.Frame fraTest 
            Caption         =   "Clinical Test"
            Height          =   1455
            Left            =   4680
            TabIndex        =   60
            Top             =   1620
            Width           =   3195
            Begin VB.TextBox txtLabDesc 
               BackColor       =   &H8000000F&
               Height          =   315
               Left            =   180
               Locked          =   -1  'True
               TabIndex        =   63
               Top             =   1020
               Width           =   2835
            End
            Begin VB.ComboBox cboTest 
               Height          =   315
               Left            =   720
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Top             =   660
               Width           =   2295
            End
            Begin VB.ComboBox cboGroup 
               Height          =   315
               Left            =   720
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   300
               Width           =   2295
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               Caption         =   "Test"
               Height          =   255
               Left            =   180
               TabIndex        =   62
               Top             =   660
               Width           =   375
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               Caption         =   "Group"
               Height          =   255
               Left            =   180
               TabIndex        =   61
               Top             =   300
               Width           =   435
            End
         End
         Begin VB.TextBox txtExportName 
            Height          =   315
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   2
            Top             =   960
            Width           =   2655
         End
         Begin VB.Frame fraCase 
            Caption         =   "Case"
            Height          =   1455
            Left            =   4680
            TabIndex        =   20
            Top             =   1620
            Width           =   3195
            Begin VB.OptionButton optCase 
               Caption         =   "Lower Case"
               Height          =   255
               Index           =   2
               Left            =   180
               TabIndex        =   18
               Top             =   960
               Width           =   1215
            End
            Begin VB.OptionButton optCase 
               Caption         =   "Upper Case"
               Height          =   195
               Index           =   1
               Left            =   180
               TabIndex        =   17
               Top             =   660
               Width           =   1215
            End
            Begin VB.OptionButton optCase 
               Caption         =   "As Entered"
               Height          =   195
               Index           =   0
               Left            =   180
               TabIndex        =   16
               Top             =   360
               Value           =   -1  'True
               Width           =   1215
            End
         End
         Begin VB.CommandButton cmdDerivationExpression 
            Height          =   675
            Left            =   7200
            Picture         =   "frmDataDefinition.frx":07BC
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   3120
            Width           =   675
         End
         Begin MSMask.MaskEdBox txtDataItemLength 
            Height          =   315
            Left            =   1800
            TabIndex        =   7
            Top             =   2760
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            Enabled         =   0   'False
            MaxLength       =   5
            Format          =   "0"
            PromptChar      =   "_"
         End
         Begin VB.Frame fraCopiedFrom 
            Caption         =   "Copied from"
            Height          =   1095
            Left            =   4680
            TabIndex        =   19
            Top             =   240
            Width           =   3195
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Data item code"
               Height          =   195
               Left            =   240
               TabIndex        =   21
               Top             =   1440
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.Label lblVersionFrom 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Version"
               Height          =   195
               Left            =   285
               TabIndex        =   29
               Top             =   720
               Visible         =   0   'False
               Width           =   825
            End
            Begin VB.Label lblStudyFrom 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Study"
               Height          =   195
               Left            =   285
               TabIndex        =   30
               Top             =   360
               Width           =   765
            End
            Begin VB.Label lblCopiedFromDataItemCode 
               AutoSize        =   -1  'True
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Label4"
               Height          =   255
               Left            =   1680
               TabIndex        =   31
               Top             =   1440
               Visible         =   0   'False
               Width           =   540
            End
            Begin VB.Label lblCopiedFromVersionId 
               AutoSize        =   -1  'True
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Label3"
               Height          =   315
               Left            =   1260
               TabIndex        =   36
               Top             =   660
               Visible         =   0   'False
               Width           =   780
            End
            Begin VB.Label lblCopiedFromClinicalTrialName 
               AutoSize        =   -1  'True
               Caption         =   "Label2"
               Height          =   195
               Left            =   1260
               TabIndex        =   35
               Top             =   360
               Width           =   1560
            End
         End
         Begin VB.ComboBox cboDataType 
            Height          =   315
            ItemData        =   "frmDataDefinition.frx":0AC6
            Left            =   1800
            List            =   "frmDataDefinition.frx":0AC8
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1680
            Width           =   2655
         End
         Begin VB.TextBox txtDataItemDerivation 
            Height          =   660
            Left            =   1800
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Text            =   "frmDataDefinition.frx":0ACA
            Top             =   3120
            Width           =   5355
         End
         Begin VB.TextBox txtDataItemCode 
            Height          =   315
            Left            =   1800
            TabIndex        =   0
            Top             =   240
            Width           =   2655
         End
         Begin VB.TextBox txtDataItemHelpText 
            Height          =   735
            Left            =   1800
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Text            =   "frmDataDefinition.frx":0AE2
            Top             =   4560
            Width           =   6075
         End
         Begin VB.ComboBox cboDataItemFormat 
            Height          =   315
            Left            =   1800
            TabIndex        =   6
            Text            =   "cboDataItemFormat"
            Top             =   2400
            Width           =   2655
         End
         Begin VB.TextBox txtDataItemName 
            Height          =   315
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   1
            Text            =   "txtDataItemName"
            Top             =   600
            Width           =   2655
         End
         Begin VB.ComboBox cboUnitOfMeasurement 
            DataField       =   "UnitOfMeasurement"
            DataSource      =   "Data2"
            Height          =   315
            ItemData        =   "frmDataDefinition.frx":0AF8
            Left            =   1800
            List            =   "frmDataDefinition.frx":0AFA
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   2040
            Width           =   2655
         End
         Begin VB.Label lblReadOnly 
            Caption         =   "(eForm or Visit date)"
            Height          =   375
            Left            =   4560
            TabIndex        =   65
            Top             =   1680
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label lblDescription 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Metadata description"
            Height          =   195
            Left            =   225
            TabIndex        =   64
            Top             =   3840
            Width           =   1485
         End
         Begin VB.Label lblExportCode 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Question export code"
            Height          =   195
            Left            =   200
            TabIndex        =   37
            Top             =   960
            Width           =   1515
         End
         Begin VB.Label lblDataItemType 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Question type"
            Height          =   195
            Left            =   750
            TabIndex        =   34
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label lblDerivation 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Derivation"
            Height          =   195
            Left            =   960
            TabIndex        =   33
            Top             =   3120
            Width           =   720
         End
         Begin VB.Label lblDataItemCode 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Question code"
            Height          =   195
            Left            =   645
            TabIndex        =   32
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label lblDataItemHelpText 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "User help text"
            Height          =   195
            Left            =   720
            TabIndex        =   28
            Top             =   4560
            Width           =   975
         End
         Begin VB.Label lblDataItemLength 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Maximum length"
            Enabled         =   0   'False
            Height          =   195
            Left            =   300
            TabIndex        =   27
            Top             =   2760
            Width           =   1380
         End
         Begin VB.Label lblDataItemFormat 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Format"
            Height          =   195
            Left            =   1200
            TabIndex        =   26
            Top             =   2400
            Width           =   480
         End
         Begin VB.Label lblDataItemName 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Question name"
            Height          =   195
            Left            =   645
            TabIndex        =   25
            Top             =   600
            Width           =   1065
         End
         Begin VB.Label lblUnitOfMeasurement 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Unit of measurement"
            Height          =   195
            Left            =   240
            TabIndex        =   24
            Top             =   2040
            Width           =   1455
         End
      End
      Begin MSComctlLib.ListView lsvRequired 
         CausesValidation=   0   'False
         Height          =   4935
         Left            =   -74880
         TabIndex        =   48
         Top             =   480
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   8705
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
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
      Begin MSFlexGridLib.MSFlexGrid gridValidation 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   44
         Top             =   480
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   8705
         _Version        =   393216
         Cols            =   4
         ScrollBars      =   2
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid gridCategories 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   42
         Top             =   480
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   8705
         _Version        =   393216
         Cols            =   4
         ScrollBars      =   2
         AllowUserResizing=   1
      End
      Begin VB.TextBox txtAfterValidationTextTabOrder 
         Height          =   615
         Left            =   -74160
         TabIndex        =   39
         Text            =   "txtAfterValidationTextTabOrder"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtAfterCategoryTextTabOrder 
         Height          =   645
         Left            =   -74160
         TabIndex        =   46
         Text            =   "txtAfterCategoryTextTabOrder"
         Top             =   2880
         Width           =   1935
      End
      Begin VB.TextBox txtAfterCategroyGridTabOrder 
         Height          =   615
         Left            =   -74160
         TabIndex        =   43
         Text            =   "txtAfterCategroyGridTabOrder"
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Frame Frame2 
         Caption         =   "Category Loading"
         Height          =   4815
         Left            =   -69840
         TabIndex        =   49
         Top             =   480
         Width           =   2895
         Begin VB.ListBox lstConnectionType 
            Height          =   1035
            Left            =   240
            TabIndex        =   59
            Top             =   1800
            Width           =   2415
         End
         Begin VB.Frame fmeLoadOptions 
            Caption         =   "Load:"
            Height          =   1335
            Left            =   240
            TabIndex        =   53
            Top             =   360
            Width           =   2415
            Begin VB.OptionButton optLoadOption 
               Caption         =   "None if at least one record is invalid"
               Height          =   435
               Index           =   2
               Left            =   120
               TabIndex        =   56
               Top             =   840
               Width           =   2175
            End
            Begin VB.OptionButton optLoadOption 
               Caption         =   "All records"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   55
               Top             =   240
               Width           =   2175
            End
            Begin VB.OptionButton optLoadOption 
               Caption         =   "Valid records only"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   54
               Top             =   585
               Value           =   -1  'True
               Width           =   2175
            End
         End
         Begin VB.TextBox txtSQL 
            Enabled         =   0   'False
            Height          =   495
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   52
            Text            =   "frmDataDefinition.frx":0AFC
            Top             =   3720
            Width           =   2535
         End
         Begin VB.TextBox txtConnect 
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   51
            Text            =   "frmDataDefinition.frx":0B1F
            Top             =   3120
            Width           =   2535
         End
         Begin VB.CommandButton cmdLoad 
            Caption         =   "&Load"
            Height          =   375
            Left            =   1920
            TabIndex        =   50
            Top             =   4320
            Width           =   855
         End
         Begin VB.Label lblSQL 
            Caption         =   "Data Source:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   58
            Top             =   3480
            Width           =   1335
         End
         Begin VB.Label lblConnect 
            Caption         =   "Connection String:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   57
            Top             =   2880
            Width           =   1575
         End
      End
   End
   Begin VB.TextBox txtAfterValidationICBOTabOrder 
      Height          =   615
      Left            =   840
      TabIndex        =   41
      Text            =   "txtAfterValidationICBOTabOrder"
      Top             =   2160
      Width           =   1935
   End
End
Attribute VB_Name = "frmDataDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1998. All Rights Reserved
'   File:       frmDataDefinition.frm
'   Author:     Andrew Newbigging June 1997
'   Purpose:    Allows data item definition to be maintained, including validation
'   and derivation rules and list of coded values for category data items.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
' WillC removed Old comments 17/8/00
'       Added validation to txtDataItemLength_LostFocus to check for maximum value
'   41  Paul Norris             29/7/99     SR's: 476,408,898,594,996,811,14921303,453,1393,446,1175
'       Amended UpdateDataItem(), SaveChangedData() and ShowDataDefinition()
'       Upgraded validation and categories tabs to display editable grid
'   42  Paul Norris             18/08/99    Amended ValidateDataFormat()
'   43  Paul Norris             18/08/99    Amended HideEditingControls() to refresh grids properly
'   44  Paul Norris             02/09/99    Amended RefreshDataItemFormat(), ShowDataDefinition()
'                                           and cboDataType_Click()
'   45  Paul Norris             03/09/99    renamed ValidationAction field to ValidationActionId
'   46  PN                      10/09/99    Changed behaviour of form so that txtDataItemLength
'       and cboDataItemFormat do not validate on Lost_Focus event. They are now validated when the
'       window is closed or when a different dataitem is to be displayed
'   47  PN                      13/09/99    Field name DataItem.Case changed to DataItem.DataItemCase
'       Mo Morris   14/9/99     cmdValidationExpression_Click, cmdDerivationExpression changed.
'                               Calls to frmExpressionbuilder replaced by call to frmCUIFunctionEditor.
'                               txtDataItemDerivation_Lostfocus has had frmExpressionBuilder
'                               replaced by frmCUIFunctionEditor.
'   PN  17/09/99    Amended UpdatedataItem() and DisplayDataItem() to include library specific fields
'                   Required and TrialTypeId
'   PN  20/09/99    Added moFormIdleWatch object to handle system idle timer resets
'   PN  21/09/99    Added DeleteNewDefinition() to remove definition if app is closed and
'                   new definition details are not completed by user
'   PN  24/09/99    Changed class names
'                   clsDataDefinitionValidations to clsDDValidations
'                   clsDataDefinitionValidation to clsDDValidation
'                   clsDataDefinitionCategories to clsDDCategories
'                   clsDataDefinitionCategory to clsDDCategory
'                   because prog id is too long with original name
'   PN  28/09/99    Amended HandleValidationGridClick() to prevent right mouse click flicker
'   PN  30/09/99    Amended DisplayDataDefinition() to specify full SQL statment
'   NCJ 28/10/99    Validation property IsTermValid renamed to IsConditionValid
'                   Refresh grid in txtCategoryEdit and txtValidationGridEdits LostFocus
'                   Save data item length in UpdateDataDefinition
'   WillC 10/11/99  Added the error handlers
'   Mo Morris   12/11/99    DAO to ADO conversion
'   WillC 12/11/99  Changed labels inputboxes msgboxes "question" instead of "data item"
'   NCJ 17/11/99    Rewrote checks on format strings for numbers
'                   Handle conversion of numeric format strings between "standard" and "regional"
'   NCJ 18/11/99    Fix to accept Backspace in DataItemLength field
'   NCJ 13/12/99    TrialIds and DataItemIds to Long
'   NCJ 17/12/99    SR 1768 Do not validate derivation on Cancel
'   NCJ 22/12/99    Removed unused DeleteNewDefinition
'                   SR 2290 Update CRFElement on current form if visible
'   NCJ 29/12/99    SR2420
'   WillC 5/1/00    Added the Option Buttons for Ado Connections to the default database
'                   or another Datasource tested against SQL and Access
'   Mo Morris  12/1/00
'                   Option buttons for Text,Default and ADO connections for category
'                   code retrieval replaced by lstConnectionType which now has for enteries
'                   "Text File","Default connection","ADO - Access","ADO - SQL Server"
'                   Selecting "ADO - Access" launches the new form frmAccessConnect
'   NCJ 12/1/00     Do not allow editing of Reject If conditions if trial status is open
'   NCJ 12 Jan 00   SRs679,2131,2132 Check trial status when confirming grid row deletes
'                   (see PromptDeleteGridRow, moved here from basCommon)
'   NCJ 13 Jan 00   Final tweaking of what's allowed when trial is open
'   SDM 31/01/00    Turn on and off the moFormIdleWatch for SR2073 and SR2830
'   TA 29/04/2000   subclassing removed
'   NCJ 31/5/00     Use new gsCANNOT_CONTAIN_INVALID_CHARS in error messages
'   WillC SR3741 17/8/00 spaces in the export code were causing problems.
'   WillC  SR3573   Now to check a listitems checkbox they must explicitly check the box clicking on the item
'                   text was causing erratic behaviour at run time but not when stepping through. Possible listview bug?
'   NCJ 14/9/00 - Support for LabTest type data items
'   TA  15/9/00 - Additional support for LabTest dataItems
'   NCJ/TA 25/9/00 - Debugging LabTest question editing
'   TA 04/10/2000: Laboratory, CTC Scheme, Clinical Test and Clinical Test Ids replace with Codes
'   TA 05/10/2000: SetCaseTestOption routine sorted out so that Case number and Test Codes are separate
'   Mo  15/11/2000  ReplaceCharacters call replaced by call to VB Function Replace
'   NCJ 7 Mar 2001  Changed validation on Export Code
'                   Do NOT try and do a Save when they click the Close box
'                   Do not validate fields when they click Cancel
'   ZA 02/08/01     SR 2747 New warning message is added if the user fails to define values for a category
'   Mo Morris   30/8/01 Db Audit (All aspects of DefaultCat (column 4 of gridCategories) removed)
'   Mo Morris   12/10/01 ScrollBars setting on gridValidation and gridCategories switched from Both to Vertical only,
'                   because of uncontrolable multievent activities that resulted in cels being cleared when the
'                   grid_Scroll event code was been called. THIs IS A WORAKAROUND
'   DPH 16/10/2001  Code added to hide the data entry control to fix the resize display problem
'   DPH 17/10/2001  Above changes amended to check where focus is via API call
'   DPH 23/10/2001  Throw away disallowed characters on KeyPress event
'   DPH 24/10/2001  Inactive Categories check in cmdOK_Click event
'   DPH 05/11/2001  Changed check for no defined categories in cmdOK_Click event
'   NCJ 14 Jan 02 - Be careful not to load frmCUIFunctionEditor if it's not currently loaded in cmdValidationExpression_Click()
'       (Current Build Buglist 2.2.6, Bug 2)
'   MLM 17/01/02    Current Build Buglist 2.2.3, Bug 20: When in Library mode, extra things need to be
'                   refreshed during DisplayDataDefinition in case they have been edited,
'                   including some that used to be done in Form_Load.
'   ZA 10/06/2002   Fixed bug 54 in build 2.2.13: check for validations only if blnDuringShow is set to false
'                   in UpdateValidationsWithUserEdits procedure.
' NCJ 18 Jun 02 - CBB 2.2.14/19 Preserve thousand separators in number formats
''ASH 21/06/2002 - Fixed SR 4627 and SR4640 .
'   NCJ 15 Jan 02 - Ensure we always use correct instance of frmDataList (from old 3.0)
' REM 23/07/02 - Fix CBB 2.2.10 No. 10 - Created new Function to Check if active control in the cancel button, called IsActiveControlCancel
' ZA 19/08/2002 - Added MACROOnly and Description fields on this form
' ZA 21/08/2002 - Disable question type control if it is being used for form validation
' NCJ 30 Aug 02 - Smoothed out wrinkles caused by introduction of new Arezzo Expression Editor
' MLM 16/09/02: Added label explaining why the data type control is disabled.
' ASH 31/10/2002 check for new permission added in sub Display Data Definition
' NCJ 23 Jan 03             New Function Editor
' ASH 27/2/2002             Move cmdValidationExpression along with the current cell when the form is resized
' RS 13/02/2003     SR4640  New Fix. Changed Me.DataItemName to txtDataItemCode.Text
' NCJ 13 Feb 03 - When changing to type LabTest, delete derivation and validations
' NCJ 7 Apr 03 - Don't change a category question from text box when altering categories
' RS 6 Jun 03 - BUG 1458 Fixed Validation Grid Scroll Behaviour
' RS 26 Jun 03 - BUG 1863 - Only allow SQL statements starting with 'SELECT' when loading category values
' NCJ 28 Aug 03 - BUG 1913 - Sorted out "Copied from" display
' NCJ 1 Mar 04 - Changed limit from 9 digits to 15 for integers in IsValidNumberFormat
' ic 15/06/2005 added clinical coding
'------------------------------------------------------------------------------------'
Option Explicit
Option Base 0
Option Compare Binary

' enumerators for validation controls that are not validated on
' the Lost_Focus event

Private Enum eIsDefinitionValidResult
    LengthNotValid = 0
    LengthUnableToChange = 1
    FormatNotValid = 2
    FormatUnableToChange = 3
    LabTestNotValid = 4
    AllValid = 5
    'ic 15/06/2005 clinical coding: added invalid dictionary enum
    DictionaryNotValid = 6
End Enum

' the categories grid constants
Private Const mcValueCodeCol = 1
Private Const mcValueCol = 2
Private Const mcActiveCol = 3

Private Const mcGridInvalidBackColour = vbYellow

Private mtSelectedValidationGridCell As SelectedGridCell
Private mtSelectedCategoryGridCell As SelectedGridCell

' the validation grid constants
Private Const mcValidationTypeCol = 1
Private Const mcValidationTermCol = 2
Private Const mcValidationMessageCol = 3

Private mbRefreshCategoryGrid As Boolean
Private mbRefreshValidationGrid As Boolean

Private moValidations As clsDDValidations
Private WithEvents moNewValidation As clsDDValidation
Attribute moNewValidation.VB_VarHelpID = -1
Private moCategories As clsDDCategories
Private WithEvents moNewCategory As clsDDCategory
Attribute moNewCategory.VB_VarHelpID = -1

Private mlSelCategoryCol As Long
Private mlSelCategoryRow As Long
Private mbIsNew As Boolean

Private mlDataItemId As Long
Private msDataItemName As String

Private msClinicalTrialName As String 'used to store form property
Private mlClinicalTrialId As Long 'used to store form property
Private mnVersionId As Integer 'used to store form property
Private blnNameChanged As Boolean
Private blnDataChanged As Boolean
Private blnDuringShow As Boolean

'mbForcedReFocus is used by the LostFocus/GotFocus Subs to determine when the focus
'has been re-set to a control as a result of failing its validation criteria.
'When mbForcedRefocus is true it prevents the control's Tag field being set to
'the erroneously changed text. This is all to do with detecting when a dataitem has
'been changed and blnDataChanged is set to true
Private mbForcedReFocus As Boolean

' this variable allows prevention of multiple calls to SaveChangesData
' (which raises msgboxes)
Private mbDiscardChanges As Boolean

' NCJ 6 Mar 01 - Prevent lost focus events when the user cancels
Private mbCancelled As Boolean

' variable to track if the click was a right click
Private mbIsRightMouseClick As Boolean

Private msDataItemFormat As String
Private mnDataItemLength As Integer

' NCJ 9 Feb 00 - Store result of dateformat validation
Private mnDateFormatType  As Integer

Private Const msSELECTED_TEXT = "x"
Private Const msNOT_SELECTED_TEXT = vbNullString

'collection objects for lanb test questions
Private moClinicalTests As clsClinicalTests
Private moClinicalTestGroups As clsClinTestGroups

'DPH 23/10/2001 - Disallowed characters for this form
Private msDisallowed As String

' NCJ 13 Feb 03 - Remember data type
Private menDataType As DataType

'clinical coding dictionaries
Private moDictionaries As MACROCCBS30.Dictionaries

'----------------------------------------------------------------------------------------'
Private Sub cboGroup_Click()
'----------------------------------------------------------------------------------------'
     If cboGroup.ListIndex <> -1 Then
        If moClinicalTests.PopulateCombo(cboTest, cboGroup.Text) > 0 Then
            cboTest.ListIndex = 0
        End If
        SetComboDropdownWidth Me, cboTest
    End If
End Sub


'---------------------------------------------------------------------
Private Sub cboTest_Click()
'---------------------------------------------------------------------
' test chosen - set description and units
'---------------------------------------------------------------------
    If cboTest.ListIndex <> -1 Then
        txtLabDesc.Text = moClinicalTests.Item(cboTest.Text).Description
        SetUnitsComboText moClinicalTests.Item(cboTest.Text).Unit
    Else
        txtLabDesc.Text = ""
        SetUnitsComboText ""
    End If

End Sub

'---------------------------------------------------------------------
Private Sub cmdValidationExpression_LostFocus()
'---------------------------------------------------------------------
' DPH 17/10/2001 - Hide Button once used
' NCJ 14 Jan 02 - Be careful not to load frmCUIFunctionEditor if it's not currently loaded
'       (Current Build Buglist 2.2.6, Bug 2)
' NCJ 30 Aug 02 - With the new Function Editor, the focus is initially on the text box, not the OK button
'---------------------------------------------------------------------
Dim nFocusHandle As Long
Dim lCUICmdOKHwnd As Long
Dim oForm As Form

    lCUICmdOKHwnd = -1
    ' Get the window handle of the Expression Editor's current control
    ' NCJ 14 Jan 02 - Only get the cmdOK hWnd if the form is loaded
    For Each oForm In Forms
        If oForm.Name = "frmCUIFunctionEditor" Then
'            lCUICmdOKHwnd = frmCUIFunctionEditor.cmdOK.hWnd
            ' NCJ 30 Aug 02 - In MACRO 3.0, focus will be on text box, not on OK button
'            lCUICmdOKHwnd = frmCUIFunctionEditor.txtExpression.hWnd
            ' NCJ 23 Jan 03 : Oops - seems to have changed again
            lCUICmdOKHwnd = frmCUIFunctionEditor.cmdOK.hWnd
            Exit For
        End If
    Next
    
    ' DPH 23/10/2001 - Don't hide txtValidationGridEdits if going to expression builder
    ' Check who has focus - If not Arezzo Expression builder or txtbox then hide it
    nFocusHandle = GetFocus
    If txtValidationGridEdits.Visible And ((nFocusHandle <> txtValidationGridEdits.hWnd) And (nFocusHandle <> lCUICmdOKHwnd)) Then
        txtValidationGridEdits.Visible = False
    End If
    If ((nFocusHandle <> txtValidationGridEdits.hWnd) And (nFocusHandle <> lCUICmdOKHwnd)) Then
        cmdValidationExpression.Visible = False
    End If

End Sub

'---------------------------------------------------------------------
Private Sub Form_Resize()
'---------------------------------------------------------------------

'---------------------------------------------------------------------
Dim lGridWidth As Long

    If Me.Width < 8340 Then
        Me.Width = 8340
    End If
        
'    'ASH 27/1/2003 set the width of the form so that 2 columns in the grid are always visible
'    If Me.Width < gridValidation.ColWidth(0) + gridValidation.ColWidth(1) + gridValidation.ColWidth(2) Then
'        Me.Width = gridValidation.ColWidth(0) + gridValidation.ColWidth(1) + gridValidation.ColWidth(2) + cmdValidationExpression.Width
'    End If
    
    Me.Height = 6850
    sstDataDefinition.Width = Me.Width - 135
    frmData.Width = sstDataDefinition.Width - 205
    gridValidation.Width = sstDataDefinition.Width - 270
    lsvRequired.Width = gridValidation.Width
    
    'WillC SR3573 Explicitly stated the listviews backcolor to help refreshing
    lsvRequired.BackColor = vbWindowBackground
    
    Frame2.Left = sstDataDefinition.Width - 3045
    gridCategories.Width = sstDataDefinition.Width - 3270
    SaveSetting App.Title, "QuestionDefinition", "Width", Me.Width
    
    'ASH 27/2/2002 Move cmdValidationExpression along with the current cell when the form is resized
    lGridWidth = gridValidation.ColWidth(0) + gridValidation.ColWidth(1) + gridValidation.ColWidth(2) _
        + gridValidation.ColWidth(3) + 250
    If sstDataDefinition.Width <= lGridWidth Then
    On Error Resume Next
        If Me.ActiveControl.Name = "txtValidationGridEdits" Then
            gridValidation.Width = Me.ScaleWidth
            'if second column
            If gridValidation.Col = mcValidationTermCol Then
                txtValidationGridEdits.Width = gridValidation.CellWidth
                    If cmdValidationExpression.Visible = True Then
                    cmdValidationExpression.Left = gridValidation.CellLeft + gridValidation.CellWidth - _
                    cmdValidationExpression.Width + 120
                End If
                'MOVE THE BUTTON ALONG AS THE FORM IS REDUCED IN SIZE
                If cmdValidationExpression.Left + cmdValidationExpression.Width > Me.ScaleWidth Then
                    cmdValidationExpression.Left = gridValidation.CellWidth - cmdValidationExpression.Width
                End If
                'Show the expression editor when the the col(2)is longer than the form
                If gridValidation.ColWidth(0) + gridValidation.ColWidth(1) + gridValidation.ColWidth(2) > Me.ScaleWidth Then
                    cmdValidationExpression.Left = Me.ScaleWidth - cmdValidationExpression.Width
                    txtValidationGridEdits.Width = gridValidation.ColWidth(2)
                End If
            'if third column
            ElseIf gridValidation.Col = mcValidationMessageCol Then
                If cmdValidationExpression.Visible = True Then
                    cmdValidationExpression.Left = gridValidation.CellLeft + gridValidation.CellWidth - _
                    cmdValidationExpression.Width + 120
                End If
                'MOVE THE BUTTON ALONG AS THE FORM IS REDUCED IN SIZE
                If cmdValidationExpression.Left + cmdValidationExpression.Width > Me.ScaleWidth Then
                    cmdValidationExpression.Left = Me.ScaleWidth - cmdValidationExpression.Width
                End If
                'Show the expression editor when the the col(3)is longer than the form
                If gridValidation.ColWidth(1) + gridValidation.ColWidth(2) + gridValidation.ColWidth(3) > Me.ScaleWidth Then
                    cmdValidationExpression.Left = Me.ScaleWidth - cmdValidationExpression.Width
                    txtValidationGridEdits.Width = gridValidation.ColWidth(3)
                End If
            End If
        End If
    End If
 
End Sub

'---------------------------------------------------------------------
Private Sub gridValidation_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------

    ' maintain the flag for the mouse click
    If Button = vbRightButton Then
        mbIsRightMouseClick = True
    
    Else
        mbIsRightMouseClick = False
    
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub icboValidationType_LostFocus()
'---------------------------------------------------------------------
'Added by DPH 16/10/2001 - To Hide data entry control for resize display problem
'---------------------------------------------------------------------
    Call RefreshValidationsGrid
    icboValidationType.Visible = False
End Sub

'---------------------------------------------------------------------
Private Sub lstConnectionType_Click()
'---------------------------------------------------------------------
'Added by Mo Morris 10/1/00
'---------------------------------------------------------------------

    Select Case lstConnectionType.Text
    Case "Text File"
        Call EnableODBCControls(True)
    Case "ADO - Default connection"
        txtConnect.Text = "MacroADODBConnection"
        Call EnableODBCControls(False)
    Case "ADO - Access"
        frmAccessConnect.Show vbModal
        'txtConnect.Text = ""
        Call EnableODBCControls(False)
    Case "ADO - SQL Server"
        '   ATN 27/1/2000   SR 2612
        '   Set the database type to SQL server
        frmADOConnect2.DatabaseType = MACRODatabaseType.sqlserver
        frmADOConnect2.Show vbModal
        txtConnect.Text = frmADOConnect2.ConnectString
        Call EnableODBCControls(False)
    End Select

End Sub

'---------------------------------------------------------------------
Private Sub lsvRequired_GotFocus()
'---------------------------------------------------------------------
'Explicitly stated the listviews backcolor to help refreshing so that
'it doesn't grey out anymore when switching focus between windows.
'---------------------------------------------------------------------

    'WillC SR3573 Explicitly stated the listviews backcolor to help refreshing
    lsvRequired.BackColor = vbWindowBackground
    lsvRequired.Width = gridValidation.Width
    
End Sub

'---------------------------------------------------------------------
Private Sub lsvRequired_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'---------------------------------------------------------------------
    'SDM 14/12/99 SR1177
    Item.Selected = True

End Sub

'---------------------------------------------------------------------
Private Sub lsvRequired_ItemClick(ByVal Item As MSComctlLib.ListItem)
'---------------------------------------------------------------------
    'SDM 14/12/99 SR1177
    'WillC removed the line below SR3573 causes erratic behaviour at run time
    'Item.Checked = Not Item.Checked
    
End Sub


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
Public Property Get DataItemId() As Long
'---------------------------------------------------------------------

    DataItemId = mlDataItemId

End Property

'---------------------------------------------------------------------
Public Property Let DataItemId(tmpDataItemId As Long)
'---------------------------------------------------------------------

    mlDataItemId = DataItemId

End Property

'---------------------------------------------------------------------
Public Property Get DataItemName() As String
'---------------------------------------------------------------------

    DataItemName = msDataItemName

End Property

'---------------------------------------------------------------------
Public Property Let DataItemName(DataItemName As String)
'---------------------------------------------------------------------

    msDataItemName = DataItemName

End Property

'---------------------------------------------------------------------
Private Function IsDefinitionValid() As eIsDefinitionValidResult
'ic 10/06/2005 added clinical coding
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    IsDefinitionValid = eIsDefinitionValidResult.AllValid
    
    If frmMenu.TrialStatus > 1 And cboDataItemFormat.Text < msDataItemFormat Then
        IsDefinitionValid = FormatUnableToChange
        Exit Function
    End If
    
    With cboDataType
'        If Not ValidateDataFormat(cboDataItemFormat.Text, .ItemData(.ListIndex), False) Then
        ' Display message - NCJ 17 Nov 99
        If Not ValidateDataFormat(cboDataItemFormat.Text, .ItemData(.ListIndex), True) Then
            'TA 8/6/2000 SR 3571: moved here from ValidateDataFormat function
            cboDataItemFormat.SetFocus
            IsDefinitionValid = FormatNotValid
            Exit Function
        End If
        Select Case .ItemData(.ListIndex)
        Case DataType.Text, DataType.Thesaurus
            ' this must have a max length
            If txtDataItemLength.Text = "" Or Val(txtDataItemLength) = 0 Then
            'Forced user to enter a max length to avoid type mismatch error. JL 3/09/98.Bug no 425.
                IsDefinitionValid = LengthNotValid
                Exit Function
            End If
            
            If Not gblnValidString(txtDataItemLength.Text, valNumeric) Then
                IsDefinitionValid = LengthNotValid
                Exit Function
            End If
            
            '   ATN 29/4/1999   SR823
            '   Added validation for maximum data item length
            If IsNumeric(txtDataItemLength.Text) Then
                If Val(txtDataItemLength.Text) > 255 Then
                    IsDefinitionValid = LengthNotValid
                    Exit Function
                End If
                
                If frmMenu.TrialStatus > 1 And txtDataItemLength.Text < mnDataItemLength Then
                    IsDefinitionValid = LengthUnableToChange
                    Exit Function
                End If

            Else
                IsDefinitionValid = LengthNotValid
            
            End If
            
#If CLINICALCODING Then
            'ic 10/06/2005 clinical coding: added check for no dictionary selected
            If (.ItemData(.ListIndex) = DataType.Thesaurus) Then
                If moDictionaries.Count = 0 Then
                    IsDefinitionValid = DictionaryNotValid
                End If
            End If
#End If
        
        Case DataType.LabTest
            ' Check they've selected a clinical test & group
            If cboTest.ListIndex = -1 Then
                IsDefinitionValid = LabTestNotValid
                Exit Function
            End If

        End Select
        
    End With
    
Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "IsDefinitionValid")
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
Private Sub DisplayDataDefinition(vClinicalTrialId As Long, _
                                vVersionId As Integer, _
                                vClinicalTrialName As String, _
                                vDataItemId As Long, _
                                Optional bIsNew As Boolean = False)
'---------------------------------------------------------------------
' PN 10/09/99 - split code to display data from the logic of whether to display
' a new item
' this routine handles the display of a data item definition
' MLM 17/01/02: Added refreshing of various things.
' MLM 16/09/02: Added label explaining why the data type control is disabled.
' ASH 31/10/2002 check for new permission
' ic 15/06/2005 added clinical coding
'---------------------------------------------------------------------
Dim nCount As Integer
Dim rsDataItem As ADODB.Recordset
Dim rsCopiedFrom As ADODB.Recordset
Dim sSQL As String
Dim sFormat As String
'Dim lDataType As Long
Dim nColCount As Integer

    On Error GoTo ErrHandler

    Call LockWindow(Me)
    
    cmdCancel.Enabled = True

    mbIsNew = bIsNew
    
    mtSelectedValidationGridCell.Col = 1
    mtSelectedValidationGridCell.Row = 1
    mtSelectedCategoryGridCell.Col = 1
    mtSelectedCategoryGridCell.Row = 1
    blnNameChanged = False
    blnDataChanged = False
    
    blnDuringShow = True
    Set rsDataItem = New ADODB.Recordset
    Set rsDataItem = gdsDataItem(vClinicalTrialId, _
                           vVersionId, _
                           vDataItemId)
    
    mlDataItemId = rsDataItem!DataItemId
    mlClinicalTrialId = rsDataItem!ClinicalTrialId
    mnVersionId = rsDataItem!VersionId
    
    ' NCJ 13 Feb 03 - Store as module-level variable
    menDataType = rsDataItem!DataType
    
    txtDataItemCode = rsDataItem!DataItemCode
    txtDataItemCode.Enabled = False
    
    txtExportName = rsDataItem!ExportName
    txtDataItemName = rsDataItem!DataItemName
    msDataItemName = rsDataItem!DataItemName
    txtDataItemName.Tag = rsDataItem!DataItemName
    
    'ZA 19/08/2002
    chkMACROOnly.Value = rsDataItem!MACROOnly
    txtDescription.Text = RemoveNull(rsDataItem!Description)
    
    'ASH 31/10/2002 check for new permission and lock text box accordingly
    txtDescription.Locked = Not goUser.CheckPermission(gsFnEditQuestionMetadataDescription)
    lblDescription.Enabled = Not txtDescription.Locked
        
    Me.Caption = "Question Definition - " & rsDataItem!DataItemName
    
    If rsDataItem!DataItemLength = 0 Then
        txtDataItemLength.Text = ""
    Else
        txtDataItemLength.Text = rsDataItem!DataItemLength
    End If
    
    mnDataItemLength = rsDataItem!DataItemLength
    '   SR 538  ATN 9/10/98
    '   Modified to check for ANSI padding
    ' PN change - 02/09/99
    ' refresh with all applicable formats
    Call RefreshDataItemFormat(menDataType)
    sFormat = RemoveNull(rsDataItem!DataItemFormat)
    If sFormat > "" Then
        ' Convert numeric "standard" format to regional - NCJ 17 Nov 99
        Select Case menDataType
        Case DataType.IntegerData, DataType.Real, DataType.LabTest
            sFormat = ConvertStandardToLocalNum(sFormat)
        Case Else
            ' Leave it as empty string
        End Select
        cboDataItemFormat.Text = sFormat
    End If
    msDataItemFormat = sFormat
   
    txtDataItemDerivation.Text = RemoveNull(rsDataItem!Derivation)
    
    txtDataItemHelpText.Text = RemoveNull(rsDataItem!DataItemHelpText)
    
    'MLM 17/01/02: If in Library Management, need to refresh the following as they may have changed
    If frmMenu.IsAppInLibraryMode Then
        RefreshDataType 'must be done before displaying current type
        RefreshUnitOfMeasurement
        RefreshValidationTypes
        RefreshClinicalTests
    End If
    
    ' Set the selected data item type
    For nCount = 0 To cboDataType.ListCount - 1
        If cboDataType.ItemData(nCount) = menDataType Then
            cboDataType.ListIndex = nCount
        End If
    Next
    
    
    ' Set up options for specific data type
    Call cboDataType_Click
    
    'ZA 22/08/2002 - Disable this question, if it is being used for form date validation
    'MLM 16/09/02: Added label explaining why the data type control is disabled.
    cboDataType.Enabled = True
    lblReadOnly.Visible = False
    If menDataType = DataType.Date Then
        If IsFormVisitDateQuestion(vClinicalTrialId, vVersionId, vDataItemId) Then
            cboDataType.Enabled = False
            lblReadOnly.Visible = True
        End If
    End If
    
    'NCJ 17/01/00 SR2648
    ' Fill in Case options for text
    If menDataType = DataType.Text Then
        If RTrim(rsDataItem!DataItemCase) > "" Then
            ' set the correct case option for this dataitem
            Call SetCaseTestOptions(DataType.Text, rsDataItem!DataItemCase)
        End If
    End If
    
    
    'TA 15/09/2000
    ' Fill in lab test options
    If menDataType = DataType.LabTest Then
        If Not IsNull(rsDataItem!DataItemCase) Then
            ' set the correct lab for this dataitem
            Call SetCaseTestOptions(DataType.LabTest, , rsDataItem!ClinicalTestCode)
        End If
    End If
    
#If CLINICALCODING Then
    'fill in thesaurus options
    If menDataType = DataType.Thesaurus Then
        If Not IsNull(rsDataItem!DictionaryCode) Then
            ' set the correct lab for this dataitem
            Call SetCaseTestOptions(DataType.Thesaurus, , , rsDataItem!DictionaryCode)
        End If
    End If
#End If
    
    If frmMenu.TrialStatus > eTrialStatus.InPreparation And Not mbIsNew Then
        cboDataType.Enabled = False
        cmdLoad.Enabled = False
    End If
    
    If menDataType <> DataType.LabTest Then
        'if not a lab test question sort out units
        If RTrim(rsDataItem!UnitOfMeasurement) > "" Then
            For nCount = 0 To cboUnitOfMeasurement.ListCount - 1
                If cboUnitOfMeasurement.List(nCount) = rsDataItem!UnitOfMeasurement Then
                    cboUnitOfMeasurement.ListIndex = nCount
                End If
            Next
        Else
            cboUnitOfMeasurement.ListIndex = 0
        End If
    End If
    
    If frmMenu.TrialStatus > eTrialStatus.InPreparation And Not mbIsNew Then
        cboUnitOfMeasurement.Enabled = False
    End If
    
    ' NCJ 28 Aug 03 - Only show relevant "Copied from" data (MACRO 3.0 Bug 1913)
    fraCopiedFrom.Visible = False
'    If IsNull(rsDataItem!CopiedFromClinicalTrialId) Then
'        fraCopiedFrom.Visible = False
'    Else
    If Not IsNull(rsDataItem!CopiedFromClinicalTrialId) Then
        If rsDataItem!CopiedFromClinicalTrialId > 0 Then
            sSQL = "SELECT ClinicalTrialName,DataItemCode " _
                & " FROM ClinicalTrial, DataItem " _
                & " WHERE ClinicalTrial.ClinicalTrialId = " & rsDataItem!CopiedFromClinicalTrialId _
                & " AND   ClinicalTrial.ClinicalTrialId = DataItem.ClinicalTrialId " _
                & " AND   DataItem.VersionId = " & rsDataItem!CopiedFromVersionId _
                & " AND   DataItem.DataItemId = " & rsDataItem!CopiedFromDataItemId
            Set rsCopiedFrom = New ADODB.Recordset
            rsCopiedFrom.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
                
            If Not rsCopiedFrom.EOF Then
                lblCopiedFromClinicalTrialName.Caption = rsCopiedFrom!ClinicalTrialName
                lblCopiedFromVersionId.Caption = rsDataItem!CopiedFromVersionId
                lblCopiedFromVersionId.Visible = True
                lblVersionFrom.Visible = True
                lblStudyFrom.Visible = True
                fraCopiedFrom.Visible = True
            End If
        Else
            sSQL = "SELECT DataItemCode " _
                & " FROM  DataItem " _
                & " WHERE DataItem.ClinicalTrialId = " & rsDataItem!CopiedFromClinicalTrialId _
                & " AND   DataItem.VersionId = " & rsDataItem!CopiedFromVersionId _
                & " AND   DataItem.DataItemId = " & rsDataItem!CopiedFromDataItemId
            Set rsCopiedFrom = New ADODB.Recordset
            rsCopiedFrom.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
                
            If Not rsCopiedFrom.EOF Then
                lblCopiedFromClinicalTrialName.Caption = "Library"
                lblCopiedFromVersionId.Visible = False
                lblVersionFrom.Visible = False
                lblStudyFrom.Visible = True
                fraCopiedFrom.Visible = True
            End If
        End If
        If Not rsCopiedFrom.EOF Then
            lblCopiedFromDataItemCode.Caption = rsCopiedFrom!DataItemCode
            fraCopiedFrom.Visible = True
        End If
        rsCopiedFrom.Close
        Set rsCopiedFrom = Nothing
        
    End If
    
'    sstDataDefinition.TabEnabled(1) = True  ' Validation
    sstDataDefinition.Tab = 0
    
    lsvRequired.ListItems.Clear
    
    ' NCJ 13 Feb 03 - Replaced with two calls for initialising Categories and Validations
    Call InitialiseCategories
    Call InitialiseValidations
    
'    ' PN change 41
'    ' ****
'    ' first discard any existing data
'    If Not moCategories Is Nothing Then
'        moCategories.Discard
'        Set moCategories = Nothing
'    End If
'    Set moCategories = New clsDDCategories
'
'    If Not moValidations Is Nothing Then
'        moValidations.Discard
'        Set moValidations = Nothing
'    End If
'    Set moValidations = New clsDDValidations
'
'    ' then load the data
'    moCategories.Load mlDataItemId, mlClinicalTrialId, mnVersionId
'    moValidations.Load mlDataItemId, mlClinicalTrialId, mnVersionId
'
'    ' always keep dummy item in the collection so that the user can add new categories
'    ' to the grid - it will be flagged as new but not dirty so will only save if the user
'    ' edits the row entry
'    Call AddNewCategory
'    Call AddNewValidation
'
'    ' populate the categories and validations grids from data loaded
'    Call RefreshCategoriesGrid
'    Call RefreshValidationsGrid
    
    rsDataItem.Close
    Set rsDataItem = Nothing
    
    'Retrieve grid column width settings
    For nColCount = 0 To gridValidation.Cols - 1
        gridValidation.ColWidth(nColCount) = Val(GetSetting(App.Title, _
                                                            "QuestionDefinition", _
                                                            "gridValidation_ColWidth_" & nColCount, _
                                                            2000))
        If gridValidation.ColWidth(nColCount) < 1000 Then
            gridValidation.ColWidth(nColCount) = 1000
        End If
    Next nColCount
    For nColCount = 0 To gridCategories.Cols - 1
        gridCategories.ColWidth(nColCount) = Val(GetSetting(App.Title, _
                                                            "QuestionDefinition", _
                                                            "gridCategories_ColWidth_" & nColCount, _
                                                            2000))
        If gridCategories.ColWidth(nColCount) < 1000 Then
            gridCategories.ColWidth(nColCount) = 1000
        End If
    Next nColCount

    Call UnlockWindow
'    frmDataDefinition.Show
    txtDataItemName.SelStart = 0
'    txtDataItemName.SetFocus
    blnDuringShow = False
    
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "DisplayDataDefinition")
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
Private Sub InitialiseCategories()
'---------------------------------------------------------------------
' NCJ 29/12/99 - Initialise category collections
' May be called when a data item is changed FROM category to something else
'---------------------------------------------------------------------

    ' first discard any existing data
    If Not moCategories Is Nothing Then
        moCategories.Discard
        Set moCategories = Nothing
    End If
    Set moCategories = New clsDDCategories
    
    ' Load the data
    moCategories.Load mlDataItemId, mlClinicalTrialId, mnVersionId
    
    ' always keep dummy item in the collection so that the user can add new categories
    ' to the grid - it will be flagged as new but not dirty so will only save if the user
    ' edits the row entry
    Call AddNewCategory
    
    ' populate the categories grid from data loaded
    Call RefreshCategoriesGrid

End Sub

'---------------------------------------------------------------------
Private Sub InitialiseValidations()
'---------------------------------------------------------------------
' NCJ 13 Feb 03 - Initialise validations collections
' May be called when a data item is changed TO a Lab test question
'---------------------------------------------------------------------

    ' first discard any existing data
    If Not moValidations Is Nothing Then
        moValidations.Discard
        Set moValidations = Nothing
    End If
    Set moValidations = New clsDDValidations
    
    ' Load the data
    moValidations.Load mlDataItemId, mlClinicalTrialId, mnVersionId
    
    ' always keep dummy item in the collection so that the user can add new validations
    ' to the grid - it will be flagged as new but not dirty so will only save if the user
    ' edits the row entry
    Call AddNewValidation
    
    ' populate the validations grid from data loaded
    Call RefreshValidationsGrid

End Sub

'---------------------------------------------------------------------
Public Function ShowDataDefinition(vClinicalTrialId As Long, _
                                vVersionId As Integer, _
                                vClinicalTrialName As String, _
                                vDataItemId As Long, _
                                Optional bIsNew As Boolean = False) As Boolean
'---------------------------------------------------------------------
' This functions manages the flow of saving data before displaying a
' new data definition
' all external routines that attempt to display a data defintion call this function
'
' PN change 43 -  new bIsNew parameter
' PN 10/09/99 - split code to display data from the logic of whether to display
' a new item
'---------------------------------------------------------------------
 On Error GoTo ErrHandler

    If Not mbDiscardChanges Then
        If SaveChangedData Then
            Call DisplayDataDefinition(vClinicalTrialId, vVersionId, _
                                    vClinicalTrialName, vDataItemId, bIsNew)
            ShowDataDefinition = True
            frmDataDefinition.Show
            'make sure window is at the front
            frmDataDefinition.ZOrder

        Else
            ShowDataDefinition = False
            
        End If
        
    Else
        Call DisplayDataDefinition(vClinicalTrialId, vVersionId, _
                                vClinicalTrialName, vDataItemId, bIsNew)
        ShowDataDefinition = True
        frmDataDefinition.Show

    End If
    mbDiscardChanges = False
    
Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "ShowDataDefinition")
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
Private Sub AddNewCategory()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    Set moNewCategory = Nothing
    Set moNewCategory = New clsDDCategory
    
    ' keep a reference to this category because it will
    ' raise an event when it is valid
    ' we then know when to add a new row to the grid
    With moNewCategory
        .ValueID = moCategories.GetNextUniqueId(Nothing, , EFieldIDValueID)
        .ValueOrder = moCategories.GetNextUniqueId(Nothing, , EFieldIDValueOrder)
        .IsNew = True
        .IsDirty = False
    End With
    moCategories.Add moNewCategory
    
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "AddNewCategory")
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
Private Sub AddNewValidation()
'---------------------------------------------------------------------
   
    On Error GoTo ErrHandler
    
    Set moNewValidation = Nothing
    Set moNewValidation = New clsDDValidation
    
    ' keep a reference to this validation because it will
    ' raise an event when it is valid
    ' we then know when to add a new row to the grid
    With moNewValidation
        .ValidationID = moValidations.GetNextUniqueId(Nothing)
        .IsNew = True
        .IsDirty = False
    End With
    moValidations.Add moNewValidation
    
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "AddNewValidation")
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
Private Sub cboDataItemFormat_Change()
'---------------------------------------------------------------------
   
    On Error GoTo ErrHandler
    
    If cboDataItemFormat <> vbNullString Then
        If txtDataItemLength.Enabled Then
            txtDataItemLength = Len(cboDataItemFormat)
        End If
    End If
    
    If Not blnDuringShow Then
        ' NCJ 6/3/01
        If cboDataItemFormat.Tag <> cboDataItemFormat.Text Then
            blnDataChanged = True
        End If
    End If
    
    Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cboDataItemFormat_Change")
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
Private Sub cboDataItemFormat_Click()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
 
    If cboDataItemFormat <> vbNullString Then
        If txtDataItemLength.Enabled Then
            txtDataItemLength = Len(cboDataItemFormat)
        End If
    End If
    
    Exit Sub
    
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cboDataItemFormat_Click")
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
Private Sub cboDataItemFormat_GotFocus()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If Not mbForcedReFocus Then
        'Not a re-focus so set tag
        cboDataItemFormat.Tag = cboDataItemFormat.Text
    Else
        'It was a re-focus. Tag not changed and flag re-set
        mbForcedReFocus = False
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cboDataItemFormat_GotFocus")
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
Private Sub cboDataItemFormat_LostFocus()
'---------------------------------------------------------------------
   
   On Error GoTo ErrHandler
   
  ' PN 10/09/99
    ' check that dataitem format and lentgh are valid
    If cboDataItemFormat.Tag <> cboDataItemFormat.Text Then
        blnDataChanged = True
    End If
    
    Exit Sub
    
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cboDataItemFormat_LostFocus")
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
Public Function IsValidNumberFormat(ByVal sNumFormat As String, _
                                    ByVal nDataType As Integer, _
                                    ByRef sMessage As String) As Boolean
'---------------------------------------------------------------------
' Returns TRUE if sFormat is valid number format
' nDataType should be DataType.IntegerData or DataType.Real
' If FALSE, sMessage is the error message (but this isn't used at the moment)
' NCJ 17 Nov 99 (Replaces ValidateDataFormat for numbers)
' NCJ 13/9/00 - Include DataType.LabTest too
' NCJ 1 Mar 04 - Limit on integer digits changed from 9 to 15
'---------------------------------------------------------------------
Dim n As Integer
Dim nStart As Integer
Dim sNumber As String
Dim sDecPoint As String
Dim sComma As String
Dim sChar As String
Dim bInvalid As Boolean

    On Error GoTo ErrHandler

    If Not gblnValidString(sNumFormat, valOnlySingleQuotes) Then
        sMessage = "Formats" & gsCANNOT_CONTAIN_INVALID_CHARS
        IsValidNumberFormat = False
        Exit Function
    End If
    
    ' Get rid of extraneous spaces
    sNumFormat = Trim(sNumFormat)
    
    ' Allow a minus sign at the beginning
    If Left(sNumFormat, 1) = "-" Then
        nStart = 2  ' Start at 2nd character
    Else
        nStart = 1  ' Start at 1st character
    End If
    
    ' Get regional decimal point & group characters
    sDecPoint = RegionalDecimalPointChar
    sComma = RegionalThousandSeparatorChar
    
    ' Now form a numeric string from the format
    ' and check for illegal chars on the way
    ' Integers are allowed #, 9 and ,
    ' Reals also allowed decimal point
    sNumber = ""
    bInvalid = False
    For n = nStart To Len(sNumFormat)
        sChar = Mid(sNumFormat, n, 1)
        Select Case sChar
        Case "9", "#"
            ' Treat both of these as a digit
            sNumber = sNumber & "9"
        Case sComma
            ' We can ignore commas in a number (because VB does!)
        Case sDecPoint
            If nDataType = DataType.Real Or nDataType = DataType.LabTest Then
                sNumber = sNumber & sChar
            Else
                ' Don't allow decimal points for integers
                bInvalid = True
                Exit For
            End If
        Case Else
            ' Any other character isn't allowed
            bInvalid = True
            Exit For
        End Select
    Next n
    
    ' Check length of an integer
    ' NCJ 1 Mar 04 - Changed limit from 9 to 15 because AREZZO can now cope with up to 15 digits
    If nDataType = DataType.IntegerData Then
        If Len(sNumber) > 15 Then
            sMessage = "Integer formats may not be longer than 15 digits"
            IsValidNumberFormat = False
            Exit Function
        End If
    End If
    
    ' Otherwise check what happened
    ' sNumber should be numeric
    If bInvalid Or Not IsNumeric(sNumber) Then
        sMessage = sNumFormat & vbNewLine & "is not a valid format for "
        If nDataType = DataType.IntegerData Then
            sMessage = sMessage & "an integer question"
        Else
            sMessage = sMessage & "a real number question"
        End If
        IsValidNumberFormat = False
        Exit Function
    Else
        IsValidNumberFormat = True
    End If
    
    Exit Function
    
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "IsValidNumberFormat")
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
Public Function ValidateDataFormat(ByVal vFormat As String, _
                                   ByVal vDataItemType As Integer, _
                                   Optional bShowMsg As Boolean = True) As Boolean
'---------------------------------------------------------------------
' Validate a data format based on data item type
' NCJ 13/9/00 Handling for LabTest added
'---------------------------------------------------------------------
Dim sFormat As String
Dim sMessage As String

    On Error GoTo ErrHandler

    If Not gblnValidString(vFormat, valOnlySingleQuotes) Then
        If bShowMsg Then
            MsgBox "Formats" & gsCANNOT_CONTAIN_INVALID_CHARS, _
                    vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
        End If
        ValidateDataFormat = False
        Exit Function
    End If
    
    Select Case vDataItemType
    
    ' PN
    Case DataType.Text
        ' formatting is now applied to text data types from mtm v2.0
        sFormat = Replace(vFormat, "A", vbNullString)
        sFormat = Replace(sFormat, "9", vbNullString)
        sFormat = Replace(sFormat, "a", vbNullString)
        If sFormat <> vbNullString Then
            ' this string contains some other characters
            If bShowMsg Then
                MsgBox vFormat & " is not a valid format for a text question.", vbExclamation, gsDIALOG_TITLE
            End If
            ValidateDataFormat = False
            Exit Function
        End If
           
' NCJ 17 Nov 99
' Code moved to IsValidNumberFormat
' NCJ 13/9/00 LabTest added
    Case DataType.IntegerData, DataType.Real, DataType.LabTest
        ValidateDataFormat = IsValidNumberFormat(vFormat, vDataItemType, sMessage)
        If bShowMsg And sMessage > "" Then
            MsgBox sMessage, vbExclamation, gsDIALOG_TITLE
        End If
        Exit Function

            
    Case DataType.Date
        
        'changed Mo Morris 14/1/00
        If RTrim(vFormat) = "" Then
            If bShowMsg Then
                MsgBox "A date or time must have a format string.", _
                    vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
            End If
            ValidateDataFormat = False
            Exit Function
        End If
        
        ' PN change 42 - call to ValidateDateFormat
        ' NCJ 3 Feb 00 - Use new CLM format validation and store result
        mnDateFormatType = ValidateDateFormatString(vFormat)
        ' Returns type of date format string
        Select Case mnDateFormatType
        Case eDateFormatType.InvalidFormat
            If bShowMsg Then
                MsgBox vFormat & " is not a valid format for a date or time.", _
                    vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
            End If
            ValidateDataFormat = False
            Exit Function
        Case Else
            ' It's OK
        End Select
            
    End Select
    
    ValidateDataFormat = True
    
    Exit Function

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "ValidateDataFormat")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Function

'--------------------------------------------------------------------
Private Sub cboDataType_GotFocus()
'--------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If Not mbForcedReFocus Then
        'Not a re-focus so set tag
        cboDataType.Tag = cboDataType.Text
    Else
        'It was a re-focus. Tag not changed and flag re-set
        mbForcedReFocus = False
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cboDataType_GotFocus")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
        
End Sub

'--------------------------------------------------------------------
Private Sub cboDataType_LostFocus()
'--------------------------------------------------------------------

    If cboDataType.Tag <> cboDataType.Text Then blnDataChanged = True

End Sub

'---------------------------------------------------------------------
Private Sub cboUnitOfMeasurement_GotFocus()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If Not mbForcedReFocus Then
        'Not a re-focus so set tag
        cboUnitOfMeasurement.Tag = cboUnitOfMeasurement.Text
    Else
        'It was a re-focus. Tag not changed and flag re-set
        mbForcedReFocus = False
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cboUnitOfMeasurement_GotFocus")
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
Private Sub cboUnitOfMeasurement_LostFocus()
'---------------------------------------------------------------------

    If cboUnitOfMeasurement.Tag <> cboUnitOfMeasurement.Text Then blnDataChanged = True

End Sub

'---------------------------------------------------------------------
Private Function PromptLoseExistingCategories() As VbMsgBoxResult
'---------------------------------------------------------------------
Dim sMsg As String

    On Error GoTo ErrHandler

    ' if some data exists in the grid
    If moCategories.CountValid > 0 Then
        sMsg = "This action will clear the current category codes/values. Do you wish to continue?"
        PromptLoseExistingCategories = MsgBox(sMsg, vbYesNo + vbQuestion, gsDIALOG_TITLE)
    Else
        PromptLoseExistingCategories = vbYes
    End If
        
Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "PromptLoseExistingCategories")
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
Private Sub cmdValidationExpression_Click()
'---------------------------------------------------------------------
' Display the AREZZO function editor
'---------------------------------------------------------------------
Dim sType As String
Dim sTerm As String

    On Error GoTo ErrHandler

    ' NCJ 28 Jan 03 - Distinguish between Validation Condition and Validation Message
    If gridValidation.Col = mcValidationTermCol Then
        sType = "Condition"
        sTerm = "Validation condition"
    Else
        sType = "Expression"
        sTerm = "Validation message"
    End If
    frmCUIFunctionEditor.Initialise txtValidationGridEdits, _
                        sType, sTerm & " for question " & Me.DataItemName
    gridValidation.Redraw = True
    DoEvents
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdValidationExpression_Click")
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
Private Sub Form_Unload(Cancel As Integer)
'---------------------------------------------------------------------
' PN 20/09/99 - created
' This gets called if the user clicks the "close" box on the form
' NCJ 17/5/00 - We don't want to unload - instead we want to hide ourselves
' (but only if MACRO isn't exiting)
' NCJ 7/3/01 - Don't try and Save here because it can cause untold damage
'   if there's an invalid field (problems with LostFocus events and things...)
'---------------------------------------------------------------------
Dim sMsg As String

    If Me.Visible Then
        ' See if they've changed anything
        If HasChanges Then
            Cancel = 1      ' Prevent unloading
            sMsg = "You will lose the changes you have made to this question definition."
            sMsg = sMsg & vbCrLf & "Are you sure you want to close?"
        
            If DialogWarning(sMsg, , True) = vbOK Then
                ' Set focus to Cancel so Lost Focus events know not to do anything
                Call cmdCancel.SetFocus
                Call cmdCancel_Click
            End If
        End If
   End If
   
    
'        ' NCJ 17/5/00 SR3440 Force lost focus on current edit field
'        Call cmdOK.SetFocus
'        DoEvents
'
'        Cancel = 1      ' Prevent unloading
'
'        If SaveChangedData Then
'            ' We dealt with the save OK
'            Call HideWindow
'        Else
'            ' Leave the window where it is
'        End If
'
'    End If
        
End Sub

'---------------------------------------------------------------------
Private Sub gridCategories_Scroll()
'---------------------------------------------------------------------

    Call SetCategoriesGridText
    DoEvents
    gridCategories.Refresh
    
End Sub

'---------------------------------------------------------------------
Private Sub gridValidation_Scroll()
'---------------------------------------------------------------------
   
    ' RS 06/06/2003: BUG 1458: Simply move visible edit controls, do not execute events
    
    ' Move Text Edit Control
    If txtValidationGridEdits.Visible Then
        txtValidationGridEdits.Top = gridValidation.CellTop + gridValidation.Top
    End If
    
    ' Move Validation Expression Button
    If cmdValidationExpression.Visible Then
    With gridValidation
        If .Col = mcValidationTermCol Then
            cmdValidationExpression.Move .Left + .CellLeft + .CellWidth - cmdValidationExpression.Width, _
                              .Top + .CellTop - (cmdValidationExpression.Height - .CellHeight) / 2
            cmdValidationExpression.Visible = True
            cmdValidationExpression.ZOrder
        
        'ASH 06/01/2003 Bug SD466 adding the FX button to message column
        ElseIf .Col = mcValidationMessageCol Then
            cmdValidationExpression.Move .Left + .CellLeft + .CellWidth - cmdValidationExpression.Width, _
                              .Top + .CellTop - (cmdValidationExpression.Height - .CellHeight) / 2
            cmdValidationExpression.Visible = True
            cmdValidationExpression.ZOrder

        Else
            cmdValidationExpression.Visible = False
            DoEvents
        End If
    End With
    End If
  
    'Move ComboBox
    If icboValidationType.Visible = True Then
        icboValidationType.Top = gridValidation.CellTop + gridValidation.Top
    End If
    ' RS 06/06/2003: Disabled all original Code
'    Call SetValidationGridText
'    DoEvents
'    gridValidation.Refresh
'
'    'za 01/08/01, added redraw property to gridvalidation for drawing gridvalidation
'    'properly on the form
'    gridValidation.Redraw = True
        
End Sub

'---------------------------------------------------------------------
Private Sub icboValidationType_Click()
'---------------------------------------------------------------------
  
    On Error GoTo ErrHandler
  
    If icboValidationType.Visible Then
        Call UpdateValidationsWithUserEdits
    End If
        
        
        
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "icboValidationType_Click")
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
Private Sub SetupNextCellSelection(oGrid As MSFlexGrid)
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    If oGrid.Name = "gridValidation" Then
        Call SetValidationGridText
    Else
        Call SetCategoriesGridText
    End If
    
    Call SelectNextCell(oGrid)
    
    If oGrid.Name = "gridValidation" Then
        mtSelectedValidationGridCell.Col = gridValidation.Col
        mtSelectedValidationGridCell.Row = gridValidation.Row
    Else
        mtSelectedCategoryGridCell.Col = gridCategories.Col
        mtSelectedCategoryGridCell.Row = gridCategories.Row
    End If
'    Debug.Print "Col,Row:" & mtSelectedValidationGridCell.Col & ", " & mtSelectedValidationGridCell.Row
    
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
Private Sub icboValidationType_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    Select Case KeyAscii
    Case vbKeyReturn
        DoEvents
        txtAfterValidationICBOTabOrder.SetFocus
        DoEvents
        gridValidation.Refresh
        
    End Select
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "icboValidationType_KeyPress")
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
Private Sub cmdLoad_Click()
'---------------------------------------------------------------------
Dim iMousePointer As Integer
Dim lLoadOption As Long
Dim iIndex As Long

    On Error GoTo ErrHandler
    
    ' RS 26/06/2003: Only allow SELECT statement
    If UCase(Left(Me.txtSQL, 6)) <> "SELECT" Or InStr(1, UCase(Me.txtSQL), "DELETE FROM") > 0 Then
        ShowLoadError EBadSQL
        Exit Sub
    End If
    
    
    ' determine the load option selection
    For iIndex = 0 To optLoadOption.Count - 1
        If optLoadOption(iIndex).Value Then
            lLoadOption = iIndex
        End If
    Next iIndex
    
    If PromptLoseExistingCategories = vbYes Then
        ' load according to the selection
        iMousePointer = Screen.MousePointer
        Screen.MousePointer = vbHourglass
        'changed by Mo Morris 11/1/00
        If lstConnectionType.Text = "Text File" Then
            Call LoadCategories(lLoadOption, True)
        Else
            Call LoadCategories(lLoadOption, False)
        End If
        Screen.MousePointer = iMousePointer
        
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdLoad_Click")
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
Private Sub LoadCategories(eErrorLoadOptions As LoadOption, bLoadFromFile As Boolean)
'---------------------------------------------------------------------
Dim eResult As CategoryLoadResult
Dim oNewCategories As New clsDDCategories

    On Error GoTo ErrHandler

    Call HideEditingControls
    eResult = ENoError
    
    If bLoadFromFile Then
        ' load from a text file using the internal clsDDCategories object
        With cdImport
            .CancelError = True
            'za 13/05/02
            'allow user to select text file as well
            .Filter = "CSV files (*.csv)|*.csv|Text files(*.txt)|*.txt"
            
            On Error GoTo CancelSelected
            .ShowOpen
            On Error GoTo 0
            If .FileName <> vbNullString Then
                eResult = oNewCategories.LoadFromFile(.FileName, eErrorLoadOptions)
            End If
        End With
        
    Else
        'changed Mo Morris 14/1/00, on exit frmADOConnect does not place its
        'connection string in frmADOConnect.ConnectString (because it is used
        'for other purposes than this). So get its connection string now
        If lstConnectionType.Text = "ADO - SQL Server" Then
            txtConnect.Text = frmADOConnect2.ConnectString
        End If
        ' load from the specified source using the internal clsDDCategories object
        eResult = oNewCategories.LoadFromODBCSource(txtConnect, txtSQL, eErrorLoadOptions)
    End If
    
    Select Case eResult
    Case ETooFewFields, ECodesNotUnique, EValueBlank, ENoError
        ' a file format error or no error
        
        If eErrorLoadOptions = ELoadNoneIfErrors And _
            eResult <> ENoError Then
            ' the option to load is set to load none if an error ocurrs
            ' so show the failure message to the user
            Call ShowLoadError(eResult)
            
        Else
            ' load all records regardless of error
        
            ' calling the MarkAllAsDelete method will set the deleted flag on each category
            ' the collection can then be saved to commit the delete
            ' the new categories will be added to a new collection
            ' the new collection replaces the old
            moCategories.MarkAllAsDelete
            moCategories.Save
            oNewCategories.CopyProperties moCategories
            moCategories.Discard
            Set moCategories = Nothing
            
            Set moCategories = oNewCategories
            Call AddNewCategory
            Call RefreshCategoriesGrid
            Set oNewCategories = Nothing
            
        End If

    Case EFileDoesNotExist, EBadODBCSource, EBadSQL, ExclusiveNoPassword
    
        ' data entry error
        Call ShowLoadError(eResult)
    
    End Select
    
CancelSelected:

    Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "LoadCategories")
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
Private Sub ShowLoadError(eResult As CategoryLoadResult)
'---------------------------------------------------------------------
Dim sMsg As String

    On Error GoTo ErrHandler
  
    Select Case eResult
    Case EFileDoesNotExist
        sMsg = "The selected file does not exist."
    Case EBadODBCSource
        sMsg = "The ODBC connect string is invalid."
    Case EBadSQL
        sMsg = "The sql string is invalid."
    Case ETooFewFields
        sMsg = "There are not enough fields in the source file."
    Case ECodesNotUnique
        sMsg = "The codes in the source data are not all unique."
    Case EValueBlank
        sMsg = "The source data contains blank Value fields."
    Case ExclusiveNoPassword
        sMsg = "Database is password protected or locked exclusively by another user."
    End Select
        
    If sMsg <> vbNullString Then
        MsgBox sMsg, vbCritical, gsDIALOG_TITLE
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "ShowLoadError")
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
Private Sub RefreshCategoriesGrid()
'---------------------------------------------------------------------
Dim oCat As clsDDCategory
Dim lIndex As Long
Dim iRow As Integer
Dim iCol As Integer
    
    On Error GoTo ErrHandler

    lIndex = 0
    With gridCategories
        iRow = .Row
        iCol = .Col
            
        If moCategories.Count > 0 Then
            ' this call will prevent screen flicker while the grid is populated
            Call LockWindow(Me.gridCategories)
            
            .Clear
            Call SetupCategoriesGrid
            .Rows = moCategories.CountValid + 1
            
            For Each oCat In moCategories
                If Not oCat.IsDeleted Then
                    lIndex = lIndex + 1
                    
                    'SDM 08/12/99 SR2300    Categories refreshed
                    '                       into the correct row
                    .Row = oCat.ValueOrder
                    '.Row = lIndex
                    
                    If Not oCat.IsValid And oCat.IsEmpty Then
                        .Col = mcValueCodeCol
                        .CellBackColor = mcGridInvalidBackColour
                        .Col = mcValueCol
                        .CellBackColor = mcGridInvalidBackColour
                        .Col = mcActiveCol
                        .CellBackColor = mcGridInvalidBackColour
                    End If
                    
                    .Col = mcValueCodeCol
                    .Text = oCat.ValueCode
                    .Col = mcValueCol
                    .Text = oCat.Value
                    .Col = mcActiveCol
                    .CellAlignment = flexAlignCenterCenter
                    .Text = IIf(oCat.Active = 1, msSELECTED_TEXT, msNOT_SELECTED_TEXT)
                    ' set the cross reference between grid and objects
                    .RowData(.Row) = oCat.ValueID
                End If
            Next oCat
            If .Rows - 1 > iRow Then
                .Row = iRow
            End If
            If .Cols - 1 > iCol Then
                .Col = iCol
            End If
            
            ' allow refresh grid screen
            Call UnlockWindow
        End If

    End With
        
    mbRefreshCategoryGrid = False
        
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "RefreshCategoriesGrid")
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
Private Sub RefreshValidationsGrid()
'---------------------------------------------------------------------
Dim oValidation As clsDDValidation
Dim lIndex As Long
Dim iRow As Integer
Dim iCol As Integer
Dim oComboItem As ComboItem
    
    On Error GoTo ErrHandler
    
    lIndex = 0
    With gridValidation
        iRow = .Row
        iCol = .Col
        
        .Clear
        Call SetupValidationGrid
        
        If moValidations.Count > 0 Then
            ' this call will prevent screen flicker while the grid is populated
            Call LockWindow(Me.gridValidation)
            
'            .Clear
'            Call SetupValidationGrid
            .Rows = moValidations.CountValid + 1
            
            For Each oValidation In moValidations
                If Not oValidation.IsDeleted Then
                    lIndex = lIndex + 1
                    .Row = lIndex
                    
                    '  display a coloured cell if the validation item is not valid
                    If (Not oValidation.IsValid) Or (Not oValidation.IsConditionValid) Then
                        .Col = mcValidationTypeCol
                        .CellBackColor = mcGridInvalidBackColour
                        .Col = mcValidationTermCol
                        .CellBackColor = mcGridInvalidBackColour
                        .Col = mcValidationMessageCol
                        .CellBackColor = mcGridInvalidBackColour
                    End If

                    .Col = mcValidationTypeCol
                    Call SetValidationTypeCell(oValidation.ValidationTypeID)
'                    .CellPictureAlignment = flexAlignLeftCenter
'                    .CellAlignment = flexAlignRightCenter
'
'                    ' get the combo item for this validation type id
'                    ' to return the action type
'                    Set oComboItem = icboValidationType.ComboItems(oValidation.ValidationTypeID & "K")
'                    Set .CellPicture = GetImageFromKey(oComboItem.Tag)
'                    .Text = oComboItem
                    
                    '.Text = icboValidationType.ComboItems
                    .Col = mcValidationTermCol
                    .Text = oValidation.Validation
                    .Col = mcValidationMessageCol
                    .Text = oValidation.ValidationMessage
                    
                    ' set the cross reference between grid and objects
                    .RowData(.Row) = oValidation.ValidationID
                End If
            Next oValidation
            
            If .Rows - 1 > iRow Then
                .Row = iRow
            End If
            If .Cols - 1 > iCol Then
                .Col = iCol
            End If

            ' allow refresh grid screen
            Call UnlockWindow
        End If
    End With
    
    mbRefreshValidationGrid = False
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "RefreshValidationsGrid")
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
Private Sub SetValidationTypeCell(nValidationTypeId)
'---------------------------------------------------------------------
' Assume current row and column are correct
'---------------------------------------------------------------------
Dim oComboItem As ComboItem

    With gridValidation
        .CellPictureAlignment = flexAlignLeftCenter
        .CellAlignment = flexAlignRightCenter
        
        ' get the combo item for this validation type id
        ' to return the action type
        Set oComboItem = icboValidationType.ComboItems(nValidationTypeId & "K")
        Set .CellPicture = GetImageFromKey(oComboItem.Tag)
        .Text = oComboItem
        .Refresh
        'icboValidationType.Text = .Text
        'icboValidationType.ComboItems (nValidationTypeId & "K")
        Set icboValidationType.SelectedItem = icboValidationType.ComboItems(nValidationTypeId & "K")
        
    End With
    
End Sub

'---------------------------------------------------------------------
Private Sub cmdOK_Click()
'---------------------------------------------------------------------
' this routine will validate changes then save, clean up and exit
' the window is not unloaded but hidden
' REVISIONS
' DPH 24/10/2001 - Check for no active categories
' DPH 05/11/2001 - Changed check for no defined categories
' ic 14/06/2005 added clinical coding
'---------------------------------------------------------------------
Dim sMsg As String
Dim eValidResult As eIsDefinitionValidResult
Dim nColCount As Integer
Dim sErrorMsg As String
Dim nActiveCat As Integer
Dim oCategory As clsDDCategory

    On Error GoTo ErrHandler
    
    'WillC 24/2/2000 Changed mouse to signify that rebuild may take a moment
    
    Screen.MousePointer = vbHourglass
    
    'Save grid column width settings
    For nColCount = 0 To gridValidation.Cols - 1
        SaveSetting App.Title, _
                    "QuestionDefinition", _
                    "gridValidation_ColWidth_" & nColCount, _
                    gridValidation.ColWidth(nColCount)
    Next nColCount
    For nColCount = 0 To gridCategories.Cols - 1
        SaveSetting App.Title, _
                    "QuestionDefinition", _
                    "gridCategories_ColWidth_" & nColCount, _
                    gridCategories.ColWidth(nColCount)
    Next nColCount

    ' PN 10/09/99
    ' check that dataitem format and length are valid
    eValidResult = IsDefinitionValid
     
     If eValidResult = AllValid Then
        ' if all categories are valid in the collection
        ' and if the new category has been edited and is valid
        ' then save the data
        If moCategories.IsValid And moValidations.IsValid Then
            ' the categories are valid
            
            Call UpdateCategoryOrders
            Call HideEditingControls

            If txtDataItemLength.Text = "" And txtDataItemLength.Enabled = True Then
                txtDataItemLength_LostFocus
                Exit Sub
            End If
            'ZA 02/08/01 SR 2747- give a warning message if the question type is
            'category and there is no category defined
            ' DPH 05/11/2001 - Check if all categories deleted
            If (menDataType = DataType.Category) And (moCategories.IsAllDeletedOrEmpty) Then
                sErrorMsg = "You have not defined any category values for this question."
                
                If DialogWarning(sErrorMsg, , True) = vbOK Then
                    ' save and clean up
                    UpdateDataDefinition
                    ' save and clean up
                    Call HideWindow
                End If
            ElseIf (menDataType = DataType.Category) Then
                ' DPH 24/10/2001 - Count only active categories
                nActiveCat = 0
                For Each oCategory In moCategories
                    ' DPH 05/11/2001 - Added deleted check
                    If oCategory.Active And Not (oCategory.IsEmpty Or oCategory.IsDeleted) Then
                        nActiveCat = nActiveCat + 1
                    End If
                Next
                If nActiveCat = 0 Then
                    sErrorMsg = "You have not defined any active category values for this question."
                    
                    If DialogWarning(sErrorMsg, , True) = vbOK Then
                        ' save and clean up
                        UpdateDataDefinition
                        ' save and clean up
                        Call HideWindow
                    End If
                Else
                    ' save and clean up
                    UpdateDataDefinition
                    ' save and clean up
                    Call HideWindow
                End If
            Else
                ' save and clean up
                UpdateDataDefinition
                ' save and clean up
                Call HideWindow
            End If
            
        Else
        
            ' Let them cancel - NCJ 17/11/99
            cmdCancel.Enabled = True

            ' display a message pertaining to the problem
            If moCategories.CodeNotUnique > "" Then
                sMsg = "The category code " & vbCr & moCategories.CodeNotUnique & vbCr
                sMsg = sMsg & "is not unique or is the same as another description. " & vbCr
            
            ElseIf Not moCategories.IsValid Then
                sMsg = "The categories entered are not valid. Please check that the " & vbCr
                sMsg = sMsg & "codes and descriptions are not blank, and that codes are unique." & vbCr
            
            Else
                sMsg = "The validations entered are not valid. Please check that " & vbCr
                sMsg = sMsg & "all validation terms are valid and messages are not blank."
            End If
            MsgBox sMsg, vbCritical, gsDIALOG_TITLE
        End If
    
    Else
        sMsg = "The definition is not valid. " & vbCr
        ' display a message pertaining to the error id returned
        Select Case eValidResult
        Case LengthNotValid
            sMsg = sMsg & "Please check the question length."
'            sMsg = sMsg & vbCr & "You cannot save now."
            MsgBox sMsg, vbCritical, gsDIALOG_TITLE
        
        Case LengthUnableToChange
            sMsg = sMsg & "The question length may not be "
            sMsg = sMsg & "changed once a study has been opened."
'            sMsg = sMsg & vbCr & "You cannot save now."
            MsgBox sMsg, vbCritical, gsDIALOG_TITLE
        
        Case FormatNotValid
            ' If format invalid, message already given - NCJ 17 Nov 99
            sMsg = sMsg & "Please check the question format."
'            sMsg = sMsg & vbCr & "You cannot save now."
'            MsgBox sMsg, vbCritical, gsDIALOG_TITLE
        
        Case FormatUnableToChange
            sMsg = sMsg & "The question format may not be "
            sMsg = sMsg & "changed once a study has been opened."
'            sMsg = sMsg & vbCr & "You cannot save now."
            MsgBox sMsg, vbCritical, gsDIALOG_TITLE
        
        Case LabTestNotValid
            sMsg = sMsg & "You must select a clinical test "
            sMsg = sMsg & "for a lab test question."
'            sMsg = sMsg & vbCr & "You cannot save now."
            MsgBox sMsg, vbCritical, gsDIALOG_TITLE
          
        Case DictionaryNotValid
            sMsg = sMsg & "Because no dictionaries installed "
            sMsg = sMsg & "question cannot be of type Thesaurus."
'            sMsg = sMsg & vbCr & "You cannot save now."
            MsgBox sMsg, vbCritical, gsDIALOG_TITLE
            
        End Select
        
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdOK_Click")
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
Private Sub DestroyCollections()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    ' discard any existing data
    If Not moCategories Is Nothing Then
        moCategories.Discard
        Set moCategories = Nothing
    End If
    
    If Not moValidations Is Nothing Then
        moValidations.Discard
        Set moValidations = Nothing
    End If


    Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "DestroyCollections")
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
Private Sub RefreshDataType()
'---------------------------------------------------------------------
' this will read the macro table DataType to get the data types from the db
' it uses the ids read fromthe db also to populate the combo
'
' PN 10/09/99
' updated combo to read contents from table
' so here it must use the ids in the itemdata property of the combo
' not the text of the combo
' NCJ 14/9/2000 - Fill in combo "manually" rather than from DataType table
' ic 10/06/2005 added clinical coding
'---------------------------------------------------------------------
Dim rsTestCount As ADODB.Recordset

    On Error GoTo ErrHandler
    
    With cboDataType
        .Clear ' remove any existing items
        .AddItem ("Text")
        .ItemData(.NewIndex) = DataType.Text
        .AddItem ("Category")
        .ItemData(.NewIndex) = DataType.Category
        .AddItem ("Integer Number")
        .ItemData(.NewIndex) = DataType.IntegerData
        .AddItem ("Real Number")
        .ItemData(.NewIndex) = DataType.Real
        .AddItem ("Date/Time")
        .ItemData(.NewIndex) = DataType.Date
        .AddItem ("Multimedia")
        .ItemData(.NewIndex) = DataType.Multimedia
        
        'TA 27/09/2000: only allow LabTest types if clinicaltests exist
        Set rsTestCount = New ADODB.Recordset
        rsTestCount.Open "SELECT COUNT(*) FROM ClinicalTest", MacroADODBConnection
        If rsTestCount.Fields(0).Value > 0 Then
            'there is at least one row in ClinicalTest table
            .AddItem ("Laboratory Test")
            .ItemData(.NewIndex) = DataType.LabTest
        End If
        
        
#If (CLINICALCODING) Then
        'ic 10/06/2005 clinical coding: added thesaurus type to datatype combo
        .AddItem ("Thesaurus")
        .ItemData(.NewIndex) = DataType.Thesaurus
#End If
        
    End With
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "RefreshDataType")
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
Private Sub RefreshDataItemFormat(ByVal enDataType As DataType)
'---------------------------------------------------------------------
' Fill data format combo with predefined formats from database
' NCJ 17 Nov 99 - Convert numeric formats to regional
' NCJ 14/9/00 - Include LabTest type
' NCJ 24 Jun 03 - Ensure Real and Integer formats appear for LabTest questions
'---------------------------------------------------------------------
Dim rsStandardDataFormats As ADODB.Recordset
Dim sSQL As String
Dim sFormat As String

    On Error GoTo ErrHandler
    
    sSQL = "SELECT * FROM StandardDataFormat WHERE "
    If enDataType = DataType.LabTest Then
        ' Pick up Integer or Real types
        sSQL = sSQL & "(DataTypeId = " & DataType.IntegerData & _
                    " OR DataTypeId = " & DataType.Real & ")"
    Else
        ' Just pick up the relevant ones
        sSQL = sSQL & "DataTypeId = " & enDataType
    End If
    Set rsStandardDataFormats = New ADODB.Recordset
    rsStandardDataFormats.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    With cboDataItemFormat
        .Clear ' remove any existing items
        .AddItem gsEMPTY_STRING
        .ItemData(.NewIndex) = -1
        

        While Not rsStandardDataFormats.EOF
            ' Convert numeric formats to "regional" format - NCJ 17/11/99
            Select Case enDataType
            Case DataType.IntegerData, DataType.Real, DataType.LabTest
                sFormat = ConvertStandardToLocalNum(rsStandardDataFormats.Fields("DataFormat"))
            Case Else
                sFormat = rsStandardDataFormats.Fields("DataFormat")
            End Select
            .AddItem sFormat
            rsStandardDataFormats.MoveNext
        Wend
    End With
    SetComboDropdownWidth Me, cboDataItemFormat 'SDM 01/02/00 SR2840
    rsStandardDataFormats.Close
    Set rsStandardDataFormats = Nothing
    
    Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "RefreshDataItemFormat")
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
Public Sub RefreshUnitOfMeasurement()
'---------------------------------------------------------------------
Dim sCurrentSelection As String
Dim rsTemp As ADODB.Recordset
Dim sSQL As String

    On Error GoTo ErrHandler
    
    sSQL = "SELECT * FROM Units"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If cboUnitOfMeasurement.Text <> gsEMPTY_STRING Then
        sCurrentSelection = cboUnitOfMeasurement.Text
    End If
    
    cboUnitOfMeasurement.Clear ' remove any existing items
    
    cboUnitOfMeasurement.AddItem gsEMPTY_STRING
    cboUnitOfMeasurement.ItemData(cboUnitOfMeasurement.NewIndex) = -1
    
    While Not rsTemp.EOF
        cboUnitOfMeasurement.AddItem rsTemp!Unit
        If sCurrentSelection = rsTemp!Unit Then
            cboUnitOfMeasurement.Text = rsTemp!Unit
        End If
        rsTemp.MoveNext
    Wend
    
    rsTemp.Close
    Set rsTemp = Nothing
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "RefreshUnitOfMeasurement")
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
Private Sub cboDataType_Click()
'---------------------------------------------------------------------
' User is changing data type
' NCJ 13 Feb 03 - If changing to Lab test, delete validations (SR 5186)
' ic 10/06/2005 added clinical coding
'---------------------------------------------------------------------
Dim sMsg As String
Dim sAnd As String

    On Error GoTo ErrHandler

    If Not blnDuringShow Then
        ' NCJ 13 Feb 03 - Check they're really changing something
        If menDataType = cboDataType.ItemData(cboDataType.ListIndex) Then Exit Sub
        
        ' If they're changing from Category, remind them
        ' that all their categories will be deleted
        If menDataType = DataType.Category And moCategories.CountReal > 0 _
          And cboDataType.ItemData(cboDataType.ListIndex) <> DataType.Category Then
            sMsg = "Changing the data type from Category to " & cboDataType.List(cboDataType.ListIndex) & vbCrLf
            sMsg = sMsg & "means that existing category values will be deleted." & vbCrLf
            sMsg = sMsg & "Are you sure you wish to continue?"
            If DialogQuestion(sMsg) = vbYes Then
                moCategories.MarkAllAsDelete
                ' NCJ 7 Dec 99 - Must save at this point to avoid upsetting Arezzo
                ' with new (incompatible) data type
                ' NB They won't come back even if the user subsequently clicks the Cancel button!
                moCategories.Save
                ' NCJ 29/12/99 - Must clear the categories grid
                ' (so the old categories aren't there if they switch back to Category)
                Call InitialiseCategories
            Else
                ' They don't want to change the data type after all
                cboDataType.ListIndex = menDataType
                Exit Sub
            End If
        End If
        
        ' NCJ 13 Feb 03, SR 5186 - If they're changing to LabTest, remind them
        ' that all their validations will be deleted
#If CLINICALCODING Then
        'ic 12/07/2005 added clinical coding
        If (cboDataType.ItemData(cboDataType.ListIndex) = DataType.LabTest _
         Or cboDataType.ItemData(cboDataType.ListIndex) = DataType.Thesaurus) _
         And ((moValidations.CountReal > 0) Or (txtDataItemDerivation.Text > "")) Then
#Else
        If cboDataType.ItemData(cboDataType.ListIndex) = DataType.LabTest _
         And ((moValidations.CountReal > 0) Or (txtDataItemDerivation.Text > "")) Then
#End If
            sAnd = ""
            sMsg = "Changing the data type to " & cboDataType.List(cboDataType.ListIndex) & vbCrLf
            sMsg = sMsg & "means that this question's "
            If moValidations.CountReal > 0 Then
                sMsg = sMsg & "validations "
                sAnd = "and "
            End If
            If txtDataItemDerivation.Text > "" Then
                sMsg = sMsg & sAnd & "derivation "
            End If
            sMsg = sMsg & "will be deleted." & vbCrLf
            sMsg = sMsg & "Are you sure you wish to continue?"
            If DialogQuestion(sMsg) = vbYes Then
                ' Delete all the validations
                ' NB They won't come back even if the user subsequently clicks the Cancel button!
                moValidations.MarkAllAsDelete
                moValidations.Save
                ' Must clear the validations grid
                ' (so the old validations aren't there if they switch back to another data type)
                Call InitialiseValidations
                ' Clear the derivation
                txtDataItemDerivation.Text = ""
            Else
                ' They don't want to change the data type after all
                cboDataType.ListIndex = menDataType
                Exit Sub
            End If
        End If
    
    End If
    
    ' NCJ 13 Feb 03 - Remember the selected data type
    menDataType = cboDataType.ItemData(cboDataType.ListIndex)
    
    'default settings
    Call SetCaseTestOptions(menDataType, 0, "")
    
    ' PN 10/09/99
    ' updated combo to read contents from table
    ' so here it must use the ids in the itemdata property of the combo
    ' not the text of the combo
    Select Case menDataType
    Case DataType.Text
        txtDataItemLength.Enabled = True
        lblDataItemLength.Enabled = True
        
        cboUnitOfMeasurement.Enabled = False
        cboUnitOfMeasurement.ListIndex = -1
        lblUnitOfMeasurement.Enabled = False
        cboDataItemFormat.Enabled = True
        lblDataItemFormat.Enabled = True
                
        ' Derivations allowed
        txtDataItemDerivation.Enabled = True
        cmdDerivationExpression.Enabled = True
        lblDerivation.Enabled = True
        
        sstDataDefinition.TabEnabled(2) = False     ' Category
        sstDataDefinition.TabEnabled(1) = True     ' Validation
    
        'SDM 20/01/00 SR2770
        If Len(txtDataItemLength.Text) = 0 Then
            txtDataItemLength.Text = "100"
        End If
    
    Case DataType.Category
        If sstDataDefinition.TabEnabled(2) = False _
        Or blnDuringShow Then
            txtDataItemLength.Enabled = False
            lblDataItemLength.Enabled = False
            cboUnitOfMeasurement.Enabled = False
            cboUnitOfMeasurement.ListIndex = -1
            lblUnitOfMeasurement.Enabled = False
    
            cboDataItemFormat.Enabled = False
            cboDataItemFormat.Text = gsEMPTY_STRING
            lblDataItemFormat.Enabled = False
            sstDataDefinition.TabEnabled(2) = True      ' Category
            sstDataDefinition.TabEnabled(1) = True     ' Validation
            txtDataItemLength.Enabled = False
            
        End If
        txtDataItemLength.Enabled = False
        lblDataItemLength.Enabled = False
        
        ' Derivations allowed
        txtDataItemDerivation.Enabled = True
        cmdDerivationExpression.Enabled = True
        lblDerivation.Enabled = True
                
    Case DataType.IntegerData
        
        lblDataItemLength.Enabled = False
        txtDataItemLength.Enabled = False
        
        cboUnitOfMeasurement.Enabled = True
        
        'TA 28/09/2000: reset units combo
        SetUnitsComboText ""
        
        lblUnitOfMeasurement.Enabled = True
        cboDataItemFormat.Enabled = True
        lblDataItemFormat.Enabled = True
        If cboDataItemFormat = "" Then
            cboDataItemFormat.Text = "9"
        End If
        
        ' Derivations allowed
        txtDataItemDerivation.Enabled = True
        cmdDerivationExpression.Enabled = True
        
        sstDataDefinition.TabEnabled(2) = False     ' Category
        sstDataDefinition.TabEnabled(1) = True     ' Validation
        
    Case DataType.Real
        
        lblDataItemLength.Enabled = False
        txtDataItemLength.Enabled = False
        cboUnitOfMeasurement.Enabled = True
        lblUnitOfMeasurement.Enabled = True
        cboDataItemFormat.Enabled = True
        If cboDataItemFormat = "" Then
            cboDataItemFormat.Text = "9.9"
        End If
        lblDataItemFormat.Enabled = True
        
        ' Derivations allowed
        txtDataItemDerivation.Enabled = True
        cmdDerivationExpression.Enabled = True
        lblDerivation.Enabled = True
        
        sstDataDefinition.TabEnabled(2) = False     ' Category
        sstDataDefinition.TabEnabled(1) = True     ' Validation
        
    Case DataType.Date

        lblDataItemLength.Enabled = False
        txtDataItemLength.Enabled = False
        
        cboUnitOfMeasurement.Enabled = False
        cboUnitOfMeasurement.ListIndex = -1
        lblUnitOfMeasurement.Enabled = False
    
        cboDataItemFormat.Enabled = True
        lblDataItemFormat.Enabled = True
        
        ' Derivations allowed
        txtDataItemDerivation.Enabled = True
        cmdDerivationExpression.Enabled = True
        lblDerivation.Enabled = True
        
        sstDataDefinition.TabEnabled(2) = False     ' Category
        sstDataDefinition.TabEnabled(1) = True     ' Validation
    
    Case DataType.Multimedia

        lblDataItemLength.Enabled = False
        txtDataItemLength.Enabled = False
        
        cboUnitOfMeasurement.Enabled = False
        lblUnitOfMeasurement.Enabled = False
        cboDataItemFormat.Enabled = False
        lblDataItemFormat.Enabled = False
        
        ' Derivations not allowed
        txtDataItemDerivation.Enabled = False
        cmdDerivationExpression.Enabled = False
        lblDerivation.Enabled = False
        
        sstDataDefinition.TabEnabled(2) = False     ' Category
        sstDataDefinition.TabEnabled(1) = True     ' Validation
     
    ' NCJ 14/9/00 - Deal with LabTest data type
    ' Disable derivation and validation
    Case DataType.LabTest
        
        'set labtest to the default( top) one
        Call SetCaseTestOptions(eDataType.LabTest, , "")
        
        lblDataItemLength.Enabled = False
        txtDataItemLength.Enabled = False
    
        ' Derivations not allowed
        txtDataItemDerivation.Enabled = False
        cmdDerivationExpression.Enabled = False
        lblDerivation.Enabled = False
    
        'enable format combo
        cboDataItemFormat.Enabled = True
        
        'za 02/08/01, fixed 4299 bug
        'just set the visibility to true and call the refresh method
        cboDataItemFormat.Visible = True
        cboDataItemFormat.Refresh
        
        lblDataItemFormat.Enabled = True
    
        ' Disable units combo - but will be filled in when they select Clinical Test
        cboUnitOfMeasurement.Enabled = False
        lblUnitOfMeasurement.Enabled = False
    
        sstDataDefinition.TabEnabled(2) = False     ' Category
        sstDataDefinition.TabEnabled(1) = False     ' No Validation allowed
    
    Case DataType.Thesaurus
        'ic 10/06/2005 clinical coding: added thesaurus handler
        Call SetCaseTestOptions(eDataType.Thesaurus, , "")
        
        txtDataItemLength.Enabled = True
        lblDataItemLength.Enabled = True
        
        cboUnitOfMeasurement.Enabled = False
        cboUnitOfMeasurement.ListIndex = -1
        lblUnitOfMeasurement.Enabled = False
        cboDataItemFormat.Enabled = True
        lblDataItemFormat.Enabled = True
                
        'derivations not allowed
        txtDataItemDerivation.Enabled = False
        cmdDerivationExpression.Enabled = False
        lblDerivation.Enabled = True
        
        sstDataDefinition.TabEnabled(2) = False     ' Category
        sstDataDefinition.TabEnabled(1) = False     ' Validation
    
        If Len(txtDataItemLength.Text) = 0 Then
            txtDataItemLength.Text = "100"
        End If
    
    End Select
    
    If Not blnDuringShow Then
        ' PN change 44 - 02/09/99
        ' Refresh data type with appropriate formats only
        Call RefreshDataItemFormat(menDataType)
        
    End If
    
    If cboDataType.Tag <> cboDataType.Text And Not blnDuringShow Then
        blnDataChanged = True
    End If
    
    If Not txtDataItemLength.Enabled Then
        txtDataItemLength = vbNullString
    End If
    
   
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cboDataType_Click")
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
' MLM 17/01/02: Moved code for displaying validation and clinical test dropdowns into new subs
'   RefreshValidationTypes and RefreshClinicalTests.
' ic 14/06/2005 added clinical coding
'---------------------------------------------------------------------
'Dim oComboItem As ComboItem
'Dim rsValidationTypes As ADODB.Recordset
'Dim sSQL As String
    
    On Error GoTo ErrHandler
    
    'changed Mo Morris 10/1/00
    lstConnectionType.AddItem "Text File"
    lstConnectionType.AddItem "ADO - Default connection"
    lstConnectionType.AddItem "ADO - Access"
    lstConnectionType.AddItem "ADO - SQL Server"
    
    ' Turn on key preview for form, so that F1 (Help) can be trapped by form
    Me.KeyPreview = True
    
    Me.Icon = frmMenu.Icon
    
    Me.Top = frmMenu.tlbMenu.Height
    
    Me.Height = 6850
    Me.Width = Val(GetSetting(App.Title, "QuestionDefinition", "Width", 8340))
    Me.Left = (frmMenu.Width - Me.Width) / 2

    RefreshDataType
    RefreshUnitOfMeasurement
    blnDataChanged = False
    mbForcedReFocus = False
    cmdCancel.Enabled = True
    
    ' load the image list that the imagecombo uses
    With ilValidationTypes
        .ListImages.Add , "0K", LoadResPicture(gsIconMandatoryValidation, vbResIcon)
        .ListImages.Add , "1K", LoadResPicture(gsIconWarningValidation, vbResIcon)
        .ListImages.Add , "2K", LoadResPicture(gsIconDataManagerValidation, vbResIcon)
    End With

'    ' read all records in the table
'    sSQL = "select ValidationTypeId" & _
'        ",ValidationActionId,ValidationTypeName from ValidationType"
'    Set rsValidationTypes = New ADODB.Recordset
'    rsValidationTypes.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
'    With rsValidationTypes
'        Set icboValidationType.ImageList = ilValidationTypes
'        Do While Not .EOF
'            ' use the ValadationTypeAction as the unique Key
'
'            ' PN change - 03/09/99
'            ' renamed ValidationAction to ValidationActionId
'            Set oComboItem = icboValidationType.ComboItems.Add(, _
'                    .Fields("ValidationTypeId") & "K", .Fields("ValidationTypeName"))
'            oComboItem.Image = .Fields("ValidationActionId") & "K"
'            oComboItem.Tag = .Fields("ValidationActionId") & "K"
'            .MoveNext
'        Loop
'    End With
'    rsValidationTypes.Close
'    Set rsValidationTypes = Nothing
'
'    ' dissallow a blank item to be selected in the validation type combo
'    icboValidationType.Text = icboValidationType.ComboItems(1)
   
    Call RefreshValidationTypes
   
    'TA 15/09/2000: set up clincial test data
    Set moClinicalTestGroups = New clsClinTestGroups
    Set moClinicalTests = New clsClinicalTests
    
    Call RefreshClinicalTests
    
#If CLINICALCODING Then
    'ic 14/06/2005 clinical coding: refresh dictionaries
    Set moDictionaries = New MACROCCBS30.Dictionaries
    Call RefreshDictionaries
#End If
'
'    Call moClinicalTestGroups.PopulateCombo(cboGroup)
'    SetComboDropdownWidth Me, cboGroup
'
'    If cboGroup.ListCount > 0 Then
'        cboGroup.ListIndex = 0
'        Call moClinicalTests.PopulateCombo(cboTest, cboGroup.Text)
'        SetComboDropdownWidth Me, cboTest
'        If cboTest.ListCount > 0 Then
'            cboTest.ListIndex = 0
'        End If
'    End If

    
    If frmMenu.IsAppInLibraryMode Then
        sstDataDefinition.TabVisible(3) = True
    Else
        sstDataDefinition.TabVisible(3) = False
    End If
   
   ' DPH 23/10/2001 - Set up Disallowed string
   ' 34 - " 124 - | 126 - ~
   msDisallowed = "`" & Chr(34) & Chr(124) & Chr(126)
   
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

'-----------------------------------------------------------------------------------------------------
Private Sub RefreshValidationTypes()
'-----------------------------------------------------------------------------------------------------
' MLM 17/01/02: Bug 20 from 2.2.3 Current Build Buglist.
' This code has been moved from Form_Load to here so that it can also be called from
' DisplayDataDefinition when in Library mode, so that changes to validation types are shown.
'-----------------------------------------------------------------------------------------------------
Dim oComboItem As ComboItem
Dim rsValidationTypes As ADODB.Recordset
Dim sSQL As String
    
    On Error GoTo ErrLabel
    
    'MLM 17/01/02: This is required now that this code is called repeatedly
    icboValidationType.ComboItems.Clear
    
    ' read all records in the table
    sSQL = "select ValidationTypeId" & _
        ",ValidationActionId,ValidationTypeName from ValidationType"
    Set rsValidationTypes = New ADODB.Recordset
    rsValidationTypes.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    With rsValidationTypes
        Set icboValidationType.ImageList = ilValidationTypes
        Do While Not .EOF
            ' use the ValidationTypeAction as the unique Key
            
            ' PN change - 03/09/99
            ' renamed ValidationAction to ValidationActionId
            Set oComboItem = icboValidationType.ComboItems.Add(, _
                    .Fields("ValidationTypeId") & "K", .Fields("ValidationTypeName"))
            oComboItem.Image = .Fields("ValidationActionId") & "K"
            oComboItem.Tag = .Fields("ValidationActionId") & "K"
            .MoveNext
        Loop
    End With
    rsValidationTypes.Close
    Set rsValidationTypes = Nothing

    ' disallow a blank item to be selected in the validation type combo
    icboValidationType.Text = icboValidationType.ComboItems(1)

Exit Sub

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmDataDefinition.RefreshValidationTypes"
    
End Sub

'-----------------------------------------------------------------------------------------------------
Private Sub RefreshClinicalTests()
'-----------------------------------------------------------------------------------------------------
' MLM 17/01/02: Created. Similar to RefreshValidationTypes above.
'-----------------------------------------------------------------------------------------------------

    On Error GoTo ErrLabel

    moClinicalTestGroups.Refresh
    moClinicalTests.Refresh
    
    Call moClinicalTestGroups.PopulateCombo(cboGroup)
    SetComboDropdownWidth Me, cboGroup

    If cboGroup.ListCount > 0 Then
        cboGroup.ListIndex = 0
        Call moClinicalTests.PopulateCombo(cboTest, cboGroup.Text)
        SetComboDropdownWidth Me, cboTest
        If cboTest.ListCount > 0 Then
            cboTest.ListIndex = 0
        End If
    End If

Exit Sub

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmDataDefinition.RefreshClinicalTests"
    
End Sub

#If CLINICALCODING Then
Private Sub RefreshDictionaries()
'-----------------------------------------------------------------------------------------------------
' ic 10/06/2005: clinical coding dictionaries
'-----------------------------------------------------------------------------------------------------
    On Error GoTo ErrLabel
    Dim oDictionary As MACROCCBS30.Dictionary
    Dim oDictionaryArray As Object
    Dim n As Integer

    Call moDictionaries.Init(SecurityDatabasePath)
    cboDictionary.Clear
    Set oDictionaryArray = moDictionaries.DictionaryList
    For n = 0 To moDictionaries.Count - 1
        Set oDictionary = oDictionaryArray(n)
        cboDictionary.AddItem oDictionary.Name & " v" & oDictionary.Version
        cboDictionary.ItemData(cboDictionary.NewIndex) = oDictionary.Id
    Next

Exit Sub

ErrLabel:
    Err.Raise Err.Number, , Err.Description & "|frmDataDefinition.RefreshDictionaries"
End Sub
#End If

'---------------------------------------------------------------------
Private Sub SetupCategoriesGrid()
'---------------------------------------------------------------------
   
    On Error GoTo ErrHandler
    With gridCategories
        .Row = 0
        .Col = 0
        '.ColWidth(0) = 250
        .Col = mcValueCodeCol
        '.ColWidth(mcValueCodeCol) = 500
        .Text = "Code"
        .Col = mcValueCol
        '.ColWidth(mcValueCol) = 2300
        .Text = "Description"
        .Col = mcActiveCol
        '.ColWidth(mcActiveCol) = 800
        .Text = "Active"
    End With

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "SelectedCRFFormId")
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
Private Sub SetupValidationGrid()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    With gridValidation
        .RowHeightMin = icboValidationType.Height
        .Row = 0
        .Col = mcValidationTermCol
        .Text = "Validation Term"
        '.ColWidth(mcValidationTermCol) = 2900
        .Col = mcValidationTypeCol
        .Text = "Validation Type"
        '.ColWidth(mcValidationTypeCol) = 1800
        .Col = mcValidationMessageCol
        .Text = "Validation Message"
        '.ColWidth(mcValidationMessageCol) = 1800
        .Row = 1
        .FocusRect = flexFocusHeavy
    End With
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "SetupValidationGrid")
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
Private Sub HideEditingControls()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    'SDM 31/01/00 SR2073, SR2830
    If icboValidationType.Visible Then

    End If
    
    txtValidationGridEdits.Visible = False
    icboValidationType.Visible = False
    cmdValidationExpression.Visible = False
    txtCategoryEdit.Visible = False
    
    ' PN change 43
    ' added refresh when clicking off grid after a new row is valid
    DoEvents
    gridCategories.Refresh
    gridValidation.Refresh
    
    'za 01/08/01 added redraw properties to the gridCategories and
    'gridValidation for dispalying them properly on the form
    gridCategories.Redraw = True
    gridValidation.Redraw = True
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "HideEditingControls")
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
Private Sub gridCategories_Click()
'---------------------------------------------------------------------
   
    On Error GoTo ErrHandler
    If Not mbIsRightMouseClick Then
        'za trap ============================
        If gridCategories.Row = 0 Then Exit Sub
        gridCategories.Redraw = True
        '====================================
        If gridCategories.Col <> gridCategories.ColSel Or gridCategories.Row <> gridCategories.RowSel Then
            ' do not edit since this is a row selection
            ' and may be followed by a delete
            Call HideEditingControls
            
        ElseIf gridCategories.Col = mcActiveCol Then
            Call ReselectActiveStatus
        Else
            Call HandleCategoryGridClick
            
        End If
        
    End If
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "gridCategories_Click")
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
Private Sub ReselectActiveStatus()
'---------------------------------------------------------------------
Dim oCategory As clsDDCategory

    On Error GoTo ErrHandler
 
    ' retrieve the correct object using the ValueID field stored
    ' in the text control tag property
    Set oCategory = moCategories.GetItemFromValueID(gridCategories.RowData(gridCategories.Row))

    'Put x on screen
    If gridCategories.TextMatrix(gridCategories.Row, mcActiveCol) = msSELECTED_TEXT Then
        gridCategories.TextMatrix(gridCategories.Row, mcActiveCol) = msNOT_SELECTED_TEXT
    Else
        gridCategories.TextMatrix(gridCategories.Row, mcActiveCol) = msSELECTED_TEXT
    End If
    
    'Update class
    oCategory.Active = IIf(gridCategories.TextMatrix(gridCategories.Row, mcActiveCol) = msSELECTED_TEXT, 1, 0)
    
    Set oCategory = Nothing

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "ReselectActiveStatus")
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
Private Sub gridValidation_Click()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    'ASH 23/12/2002  Study Definition Bug 348
    If Me.ActiveControl.Name = "icboValidationType" Then
        Exit Sub
    End If
    
    If Not mbIsRightMouseClick Then
        If gridValidation.Col <> gridValidation.ColSel Or gridValidation.Row <> gridValidation.RowSel Then
            ' do not edit since this is a row selection
            ' and may be followed by a delete
            Call HideEditingControls
            
        Else
            Call HandleValidationGridClick
            
        End If
    
    End If
    
    'ASH 27/01/2003
    Form_Resize
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "gridValidation_Click")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            End
    End Select
    
End Sub

'---------------------------------------------------------------------

Private Sub HandleCategoryGridClick()
'---------------------------------------------------------------------
Dim nNewRow As Integer
Dim nNewCol As Integer

    On Error GoTo ErrHandler
    
    ' save new selected cell
    With gridCategories
        'za 02/08/01 added redraw property with status set to true
        .Redraw = True
        nNewCol = .Col
        nNewRow = .Row
    
        ' set back to old cell position
        .Col = mtSelectedCategoryGridCell.Col
        .Row = mtSelectedCategoryGridCell.Row
        
        Call SetCategoriesGridText
        
        ' reset back to new selected cell
        .Col = nNewCol
        .Row = nNewRow
        
        ' set into edit mode
        Call ManageCategoryEdits(Asc(" "))
        DoEvents
        .Refresh
        'za 02/08/01 added redraw property with status set to true
        .Redraw = True
        mtSelectedCategoryGridCell.Col = .Col
        mtSelectedCategoryGridCell.Row = .Row
    
    End With

    
Exit Sub
ErrHandler:
    'changed by Mo Morris 12/1/00, error trapping changed to handle incorrect
    'clicks on the grid.
    Exit Sub
    
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "HandleCategoryGridClick")
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
Private Sub HandleValidationGridClick()
'---------------------------------------------------------------------
Dim nNewRow As Integer
Dim nNewCol As Integer

   
    On Error GoTo ErrHandler

    ' save new selected cell
    With gridValidation
        nNewCol = .Col
        nNewRow = .Row
        
        ' set back to old cell position
        .Col = mtSelectedValidationGridCell.Col
        .Row = mtSelectedValidationGridCell.Row
        
        Call SetValidationGridText
        
        ' reset back to new selected cell
        .Col = nNewCol
        .Row = nNewRow
        
        ' set into edit mode
        Call ValidationGridEdit(Asc(" "))
        DoEvents
        .Refresh
        
        'za 01/08/01, added redraw property to gridvalidation for propery display
        .Redraw = True
        mtSelectedValidationGridCell.Col = .Col
        mtSelectedValidationGridCell.Row = .Row
       ' Debug.Print "Col,Row:" & mtSelectedValidationGridCell.Col & ", " & mtSelectedValidationGridCell.Row

    End With

    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "HandleValidationGridClick")
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
Private Sub gridValidation_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    Call ValidationGridEdit(KeyAscii)

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "gridValidation_KeyPress")
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
Private Sub SetCategoriesGridText()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    'apply edits
    If txtCategoryEdit.Visible Then
        'SDM 08/12/99 SR2300 SR2302 The grid will be updated as the
        '                           user enters text so the next
        '                           line is not needed.
        'gridCategories = txtCategoryEdit
        
        txtCategoryEdit.Visible = False
        DoEvents
        txtCategoryEdit.Text = vbNullString
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "SetCategoriesGridText")
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
Private Sub SetValidationGridText()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
     
    If txtValidationGridEdits.Visible Then
       txtValidationGridEdits.Visible = False
       DoEvents
       txtValidationGridEdits.Text = vbNullString
    End If

    If icboValidationType.Visible Then

        If GetImageKey(icboValidationType.Text) <> vbNullString Then
            gridValidation = icboValidationType.Text
            Set gridValidation.CellPicture = GetImageFromKey(GetImageKey(icboValidationType.Text))
        End If
        
        
        icboValidationType.Visible = False
    End If

    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "SetValidationGridText")
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
Private Sub ValidationGridEdit(KeyAscii As Integer)
'---------------------------------------------------------------------
' NCJ 12/1/00 - Block editing of "Reject if" conditions if trial is open
'---------------------------------------------------------------------
Dim oValidation As clsDDValidation

    On Error GoTo ErrHandler
 
    ' NCJ 12/1/00 - Check for trial status = open for non-empty, non-new items
    Set oValidation = moValidations.GetItemFromValueID(gridValidation.RowData(gridValidation.Row))
    If frmMenu.TrialStatus > eTrialStatus.InPreparation And Not mbIsNew _
     And Not (oValidation.IsNew Or oValidation.IsEmpty) Then
        ' Do not allow editing of "Reject If" conditions
        If oValidation.ValidationTypeID = ValidationType.Mandatory Then
            Exit Sub
        End If
    End If
    
    Select Case gridValidation.Col
    Case mcValidationTermCol, mcValidationMessageCol
        ' edit a cell
        With txtValidationGridEdits
            'use correct font
            .FontName = gridValidation.FontName
            .FontSize = gridValidation.FontSize
    
           'position the edit box
            .Left = gridValidation.CellLeft + gridValidation.Left
            .Top = gridValidation.CellTop + gridValidation.Top
            .Width = gridValidation.CellWidth
            .Height = gridValidation.CellHeight
            .Visible = True
            .ZOrder
            
            Select Case KeyAscii
               Case 0 To Asc(" ")
                  txtValidationGridEdits = gridValidation
                  .SelStart = 1000
               Case Else
                  txtValidationGridEdits = Chr(KeyAscii)
                  .SelStart = 1
            End Select
            
            On Error Resume Next
            .SetFocus
            On Error GoTo 0
        End With
        
    Case mcValidationTypeCol


        With icboValidationType
            .Left = gridValidation.Left + gridValidation.CellLeft
            .Top = gridValidation.CellTop + gridValidation.Top
            .Width = gridValidation.CellWidth
            On Error Resume Next
            .ComboItems(GetImageKey(gridValidation)).Selected = True
            On Error GoTo ErrHandler
            
           
            .Visible = True
            .ZOrder
            .SetFocus
        End With
    
    End Select
    
    With gridValidation
        If .Col = mcValidationTermCol Then
            cmdValidationExpression.Move .Left + .CellLeft + .CellWidth - cmdValidationExpression.Width, _
                              .Top + .CellTop - (cmdValidationExpression.Height - .CellHeight) / 2
            cmdValidationExpression.Visible = True
            cmdValidationExpression.ZOrder
        
        'ASH 06/01/2003 Bug SD466 adding the FX button to message column
        ElseIf .Col = mcValidationMessageCol Then
            cmdValidationExpression.Move .Left + .CellLeft + .CellWidth - cmdValidationExpression.Width, _
                              .Top + .CellTop - (cmdValidationExpression.Height - .CellHeight) / 2
            cmdValidationExpression.Visible = True
            cmdValidationExpression.ZOrder

        Else
            cmdValidationExpression.Visible = False
            DoEvents
        End If
    End With
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "ValidationGridEdit")
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
Private Sub CategoryGridEdit(KeyAscii As Integer)
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
 ' edit a cell
    With txtCategoryEdit
        'use correct font
        .FontName = gridCategories.FontName
        .FontSize = gridCategories.FontSize

       'position the edit box
        .Left = gridCategories.CellLeft + gridCategories.Left
        .Top = gridCategories.CellTop + gridCategories.Top
        .Width = gridCategories.CellWidth
        .Height = gridCategories.CellHeight
        .Visible = True
        
        Select Case KeyAscii
           Case 0 To Asc(" ")
              txtCategoryEdit = gridCategories
' Debug.Print "CategoryGridEdit - Setting txtCategoryEdit = " & txtCategoryEdit
              .SelStart = 1000
           Case Else
              txtCategoryEdit = Chr(KeyAscii)
' Debug.Print "CategoryGridEdit - Setting txtCategoryEdit = " & txtCategoryEdit
              .SelStart = 1
        End Select
        
        On Error Resume Next
        .SetFocus
        On Error GoTo 0
    End With

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "SelectedCRFFormId")
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
Private Sub ManageCategoryEdits(KeyAscii As Integer)
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    mlSelCategoryCol = gridCategories.Col
    mlSelCategoryRow = gridCategories.Row
    
    If gridCategories.Tag = vbNullString Then
        ' only allow text box editing of the textural fields
        ' while not carrying out a drag
        Call CategoryGridEdit(KeyAscii)
        
    Else
        ' cancel the drag operation
        gridCategories.Drag vbCancel

    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "ManageCategoryEdits")
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
Private Sub gridCategories_KeyDown(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    If KeyCode = vbKeyDelete Then
        With gridCategories
            If .RowSel = .Row Then
        ' Debug.Print "row = " & .Row & " in gridCategories_KeyDown"
                ' only one row is selected
                If .RowData(.RowSel) <> -1 Then
                    ' the selected row is not the last row
        ' Debug.Print "calling DeleteSelectedCategories not the last row"
                    Call DeleteSelectedCategories

                End If

            Else
                ' more than one row selected
        ' Debug.Print "calling DeleteSelectedCategories more than 1 row "
                Call DeleteSelectedCategories

            End If

        End With
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "gridCategories_KeyDown")
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
Private Sub gridValidation_KeyDown(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    If KeyCode = vbKeyDelete Then
        With gridValidation
            If .RowSel = .Row Then
                ' only one row is selected
                If .RowData(.RowSel) <> -1 Then
                    ' the selected row is not the last row
                    Call DeleteSelectedValidation

                End If

            Else
                ' more than one row selected
                Call DeleteSelectedValidation

            End If

        End With
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "gridValidation_KeyDown")
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
Private Sub DeleteSelectedValidation()
'---------------------------------------------------------------------
Dim lIndex As Long
Dim lStartRow As Long
Dim lEndRow As Long

    On Error GoTo ErrHandler

    With gridValidation
        If .Row > .RowSel Then
            lStartRow = .RowSel
            lEndRow = .Row
        ElseIf .Row < .RowSel Then
            lStartRow = .Row
            lEndRow = .RowSel
        Else
            lStartRow = .RowSel
            lEndRow = .RowSel
        End If
        
        ' only allow deletion where only the last row is selected
        If Not (lStartRow = lEndRow And lStartRow = .Rows - 1) Then
            ' do not allow deleting of the last row
            If PromptDeleteGridRow(False) = vbYes Then

                For lIndex = lStartRow To lEndRow
                    If lIndex <> .Rows - 1 Then
                        moValidations.Remove .RowData(lIndex)
                    End If
                Next lIndex
                
                If mtSelectedValidationGridCell.Row > ((.Rows - 1) - (lEndRow - lStartRow)) Then
                    mtSelectedValidationGridCell.Row = (.Rows - 1) - (lEndRow - lStartRow) - 1
                End If
                Call RefreshValidationsGrid
                
            End If
        End If
    
    End With
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "DeleteSelectedValidation")
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
Private Sub DeleteSelectedCategories()
'---------------------------------------------------------------------
Dim lIndex As Long
Dim lStartRow As Long
Dim lEndRow As Long
Dim lRowsSelected As Long

    On Error GoTo ErrHandler

    With gridCategories
        ' determine the selected row indexes
        If .Row > .RowSel Then
            lStartRow = .RowSel
            lEndRow = .Row
        ElseIf .Row < .RowSel Then
            lStartRow = .Row
            lEndRow = .RowSel
        Else
            lStartRow = .RowSel
            lEndRow = .RowSel
        End If
        
        lRowsSelected = lEndRow - lStartRow + 1
        
        ' only allow deletion where only the last row is selected
        If Not (lStartRow = lEndRow And lStartRow = .Rows - 1) Then
            ' do not allow deleting of the last row
            If PromptDeleteGridRow(True) = vbYes Then

                For lIndex = lStartRow To lEndRow
                    If lIndex <> .Rows - 1 Then
                        moCategories.Remove .RowData(lIndex)
                    End If
                Next lIndex

'                If mtSelectedCategoryGridCell.Row >= ((.Rows - 1) - (lEndRow - lStartRow)) Then
'                    mtSelectedCategoryGridCell.Row = (.Rows - 1) - (lEndRow - lStartRow) - 1
'                End If

                If mtSelectedCategoryGridCell.Row > lEndRow Then
                ' selected cell is above last row to be deleted
                ' subtract no. of rows selected
                     mtSelectedCategoryGridCell.Row = mtSelectedCategoryGridCell.Row - lRowsSelected
                ElseIf mtSelectedCategoryGridCell.Row <= lEndRow And mtSelectedCategoryGridCell.Row >= lStartRow Then
                ' selected cell is in rows to be deleted
                'set it to row below
                    mtSelectedCategoryGridCell.Row = lStartRow - 1
                End If
                
                Call RefreshCategoriesGrid
                
            End If
        End If
    End With
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "DeleteSelectedCategories")
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
Private Sub gridCategories_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    If gridCategories.Col = mcActiveCol Then
        If KeyAscii = vbKeyReturn Then
            Call SetupNextCellSelection(gridCategories)
            Call gridCategories_Click
        Else
            Call ReselectActiveStatus
        End If

    Else
        Call ManageCategoryEdits(KeyAscii)
        
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "gridCategories_KeyPress")
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
Private Sub gridCategories_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
  ' set up the icon for the rordering of rows
    With gridCategories
        If .Tag = "" Then Exit Sub
        If (.MouseRow = .Rows - 1) Or (.MouseRow = 0) Then
            .DragIcon = LoadResPicture(gsIconNoDrop, vbResIcon)
        Else
            .DragIcon = LoadResPicture(gsIconCopy, vbResIcon)
        End If
    End With
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "gridCategories_DragOver")
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
Private Sub gridCategories_DragDrop(Source As VB.Control, X As Single, Y As Single)
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    With gridCategories
        If .Tag = "" Then Exit Sub
        
' Debug.Print "DragDrop MouseRow = " & .MouseRow

        ' dissallow all dragging from the first and the last row(the header row and the blank row)
        If (.MouseRow = .Rows - 1) Or (.MouseRow = 0) Then Exit Sub
        .Redraw = False
        
        ' reorder the row dragged
        .RowPosition(Val(.Tag)) = .MouseRow
        .Redraw = True
        
        Call UpdateCategoryOrders
    End With
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "gridCategories_DragDrop")
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
Private Sub UpdateCategoryOrders()
'---------------------------------------------------------------------

 ' update the category orders here
    Dim iIndex As Integer
    Dim oCat As clsDDCategory
    
    On Error GoTo ErrHandler
    
    With gridCategories
        For iIndex = 1 To .Rows - 1
            ' get the object taht relates to each row
            Set oCat = moCategories.GetItemFromValueID(.RowData(iIndex))
            
            If Not oCat Is Nothing Then
                ' set the ValueOrder property of the objects to
                ' be the same as the display order of the grid
                If oCat.IsValid Then
                    ' if the category is not valid then
                    ' setting this field may make it valid
                    ' so only set it if the cat is valid initially
                    oCat.ValueOrder = iIndex
                End If
            End If
        Next iIndex
    End With
    
    ' NCJ - If we call RefreshCategoriesGrid here then it undoes the DragDrop!!!
    RefreshCategoriesGrid
   
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "UpdateCategoryOrders")
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
Private Sub gridCategories_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------
   
   On Error GoTo ErrHandler
  ' maintain the flag for the mouse click
    If Button = vbRightButton Then
        mbIsRightMouseClick = True
    
    Else
        mbIsRightMouseClick = False
    
    End If
    
    With gridCategories
        .Tag = ""

        ' dissallow all dragging from the first and the last row(the header row and the blank row)
        If (.MouseRow = .Rows - 1) Or (.MouseRow = 0) Then Exit Sub
        If .MouseCol <> 0 Then Exit Sub
        
        ' if the grid was being edited then cancel the edits
        Call HideEditingControls
    
        .Tag = Str(.MouseRow)
        .Drag vbBeginDrag

    End With
    
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "gridCategories_MouseDown")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'--------------------------------------------------------------
Private Sub UpdateCategoriesWithUserEdits()
'--------------------------------------------------------------
' Update categories after a change to category edit field
' PN change 41
' ensure that the changes are carried accross to the categories collection
'--------------------------------------------------------------
Dim oCategory As clsDDCategory
Dim lPos As Long
Dim bIsValid As Boolean
Dim iCol As Byte
    
    ' The field validations in the Category object will raise errors
    On Error GoTo InvalidChar
    
    If Not blnDuringShow Then
        ' retrieve the correct object using the ValueID field stored
        ' in the text control tag property
        Set oCategory = moCategories.GetItemFromValueID(gridCategories.RowData(gridCategories.Row))
        bIsValid = oCategory.IsValid
        
        ' the object encapsulates the rules for a category item
        ' so let it catch the errors
        If gridCategories.Col = mcValueCodeCol Then
            ' use the function to catch input $ or £ chars
            ' since vb will implicitly cast the string to a long
            ' even with these chars
            'ASH 21/06/2002 Added Trim Function.
            oCategory.ValueCode = Trim(txtCategoryEdit.Text)
            gridCategories.Text = oCategory.ValueCode   'SDM 08/12/99 SR2300 SR2302
        ElseIf gridCategories.Col = mcValueCol Then
            oCategory.Value = Trim(txtCategoryEdit.Text)
            gridCategories.Text = oCategory.Value   'SDM 08/12/99 SR2300 SR2302
        End If
        
        If Not bIsValid And oCategory.IsValid Then
            ' the category was not valid but is now
            ' so repaint the back color of the grid row
            With gridCategories
                iCol = .Col
                .Col = 1
                .CellBackColor = vbWhite
                .Col = 2
                .CellBackColor = vbWhite
                .Col = 3
                .CellBackColor = vbWhite
                .Col = iCol
            End With
            
        End If
        
    End If
    Set oCategory = Nothing
    Exit Sub
    
InvalidChar:
    ' What they just typed wasn't valid so replace with what was there before
    ' as long as the current value isn't empty
    If txtCategoryEdit.Text <> vbNullString Then    ' Allow empty strings to stay NCJ 27/10/99
        Beep
        lPos = txtCategoryEdit.SelStart
        If gridCategories.Col = mcValueCodeCol Then
            txtCategoryEdit = oCategory.ValueCode
            gridCategories.Text = oCategory.ValueCode
        ElseIf gridCategories.Col = mcValueCol Then
            txtCategoryEdit = oCategory.Value
            gridCategories.Text = oCategory.Value
        End If
        If lPos > 0 Then lPos = lPos - 1
        txtCategoryEdit.SelStart = lPos
    End If

    
End Sub


'---------------------------------------------------------------------
Private Sub UpdateValidationsWithUserEdits()
'---------------------------------------------------------------------
' ensure that the changes are carried across to the validations collection
'---------------------------------------------------------------------
    Dim oValidation As clsDDValidation
    Dim lPos As Long
    Dim lColour As Long
    Dim iCol As Byte
    Dim iInsertPos As Integer
    Dim iValidationTypeID As Integer
    
    On Error GoTo ErrHandler
    On Error GoTo InvalidChar
    iInsertPos = txtValidationGridEdits.SelStart
    
    If Not blnDuringShow Then
        ' NCJ 12/1/00 - Corrected suspected error
        ' by replacing gridCategories with gridValidation
'        If gridCategories.RowData(gridCategories.Row) = "-1" Then
        If gridValidation.RowData(gridValidation.Row) = "-1" Then
            ' this grid row has yet to be added to the
            ' internal collection
            Set oValidation = moNewValidation
            
        Else
            ' retrieve the correct object using the ValueID field stored
            ' in the text control tag property
            Set oValidation = moValidations.GetItemFromValueID(gridValidation.RowData(gridValidation.Row))

        End If

        ' the object encapsulates the rules for a category item
        ' so let it catch the errors
        If gridValidation.Col = mcValidationTypeCol Then
            ' NCJ 12/1/00 - Do not allow setting of RejectIf conditions
            ' if trial is already open
            iValidationTypeID = Val(icboValidationType.SelectedItem.Key)
            If iValidationTypeID = ValidationType.Mandatory _
              And frmMenu.TrialStatus > eTrialStatus.InPreparation And Not mbIsNew Then
                ' Do not make assignment but return to previous value
                Call SetValidationTypeCell(oValidation.ValidationTypeID)
            Else
                oValidation.ValidationTypeID = Val(icboValidationType.SelectedItem.Key)
            End If
            
        ElseIf gridValidation.Col = mcValidationTermCol Then
            ' this must be validated
            ' the validation is carried out in the class
            oValidation.Validation = txtValidationGridEdits
        
        ElseIf gridValidation.Col = mcValidationMessageCol Then
            oValidation.ValidationMessage = txtValidationGridEdits
        
        End If

        'ZA - 10/06/2002 Execute the following code only if blnDuringShow
        'variable is set to false
        If oValidation.IsValid And oValidation.IsConditionValid Then
            ' the category was not valid but is now
            ' so repaint the back color of the grid row
            lColour = vbWhite
        Else
            lColour = mcGridInvalidBackColour
        End If
        
        With gridValidation
            iCol = .Col
            .Col = mcValidationTypeCol
            .CellBackColor = lColour
            .Col = mcValidationTermCol
            .CellBackColor = lColour
            .Col = mcValidationMessageCol
            .CellBackColor = lColour
            .Col = iCol
        End With
        
        
        If txtValidationGridEdits.Visible Then
            txtValidationGridEdits.SetFocus
            txtValidationGridEdits.SelStart = iInsertPos
        End If
        
        Set oValidation = Nothing
    
    End If
    Exit Sub
    
InvalidChar:
    Beep
    lPos = txtValidationGridEdits.SelStart
    If gridValidation.Col = mcValidationTypeCol Then
        txtValidationGridEdits = oValidation.ValidationTypeID
        gridValidation.Text = oValidation.ValidationTypeID
    ElseIf gridValidation.Col = mcValidationTermCol Then
        txtValidationGridEdits = oValidation.Validation
        gridValidation.Text = oValidation.Validation
    ElseIf gridValidation.Col = mcValidationMessageCol Then
        txtValidationGridEdits = oValidation.ValidationMessage
        gridValidation.Text = oValidation.ValidationMessage
    End If
    If lPos > 0 Then lPos = lPos - 1
    txtValidationGridEdits.SelStart = lPos
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "UpdateValidationsWithUserEdits")
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
Private Function GetImageKey(sText As String) As String
'---------------------------------------------------------------------
   On Error GoTo ErrHandler
  Dim oComboItem As ComboItem
    
    For Each oComboItem In icboValidationType.ComboItems
        If oComboItem.Text = sText Then
            GetImageKey = oComboItem.Tag
            Exit For
        End If
    Next oComboItem
    
Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "GetImageKey")
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
Private Function GetImageFromKey(sKey As String) As IPictureDisp
'---------------------------------------------------------------------
' this function will map a tag from the imagecombo to a picture
' each item in the validationtype table will map onto one of the three images
' when the imagecombo is loaded the tag is set to the image key
'---------------------------------------------------------------------
   On Error GoTo ErrHandler
  Select Case sKey

    Case "0K"
        Set GetImageFromKey = LoadResPicture(gsIconMandatoryValidation, vbResIcon)
        
    Case "1K"
        Set GetImageFromKey = LoadResPicture(gsIconWarningValidation, vbResIcon)
        
    Case "2K"
        Set GetImageFromKey = LoadResPicture(gsIconDataManagerValidation, vbResIcon)
        
    End Select
    
Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "GetImageFromKey")
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
Private Sub moNewCategory_Valid(bIsValid As Boolean)
'---------------------------------------------------------------------
   
   On Error GoTo ErrHandler
 
 If Not blnDuringShow Then
        If bIsValid Then
            ' the last item in the grid is now valid so add a new one
            ' then refresh the grid
            Call AddNewCategory
            mbRefreshCategoryGrid = True
        End If
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "moNewCategory_Valid")
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
Private Sub moNewValidation_Valid(bIsValid As Boolean)
'---------------------------------------------------------------------
  
   If Not blnDuringShow Then
        If bIsValid Then
            ' the last item in the grid is now valid so add a new one
            ' then refresh the grid
            Call AddNewValidation
            mbRefreshValidationGrid = True
        End If
    End If
    
    
End Sub

'---------------------------------------------------------------------
Private Sub optCase_Click(Index As Integer)
'---------------------------------------------------------------------

 On Error GoTo ErrHandler
    ' maintain the changed flag
    If Not blnDuringShow Then
        blnDataChanged = True
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "optCase_Click")
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
Private Sub txtAfterCategoryTextTabOrder_GotFocus()
'---------------------------------------------------------------------

   On Error GoTo ErrHandler
  ' force the grid to get focus again
    ' set the cell selected to the one after the current cell
    Call SetupNextCellSelection(gridCategories)
    If gridCategories.Col = mcActiveCol Then
        gridCategories.SetFocus
    Else
        Call gridCategories_Click
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtAfterCategoryTextTabOrder_GotFocus")
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
Private Sub txtAfterCategroyGridTabOrder_GotFocus()
'---------------------------------------------------------------------
 
    txtAfterCategoryTextTabOrder.SetFocus
     
End Sub

'---------------------------------------------------------------------
Private Sub txtAfterValidationTextTabOrder_GotFocus()
'---------------------------------------------------------------------
' this event fires when the grid loses focus
' force the grid to get focus again
' set the cell selected to the one after the current cell
'---------------------------------------------------------------------
    
    Call SetupNextCellSelection(gridValidation)
    Call gridValidation_Click
    
End Sub

'---------------------------------------------------------------------
Private Sub txtAfterValidationICBOTabOrder_GotFocus()
'---------------------------------------------------------------------
' this event fires when the grid loses focus
' force the grid to get focus again
' set the cell selected to the one after the current cell
'---------------------------------------------------------------------

    Call SetupNextCellSelection(gridValidation)
    Call gridValidation_Click
    
End Sub

'---------------------------------------------------------------------
Private Sub txtCategoryEdit_Change()
'---------------------------------------------------------------------
' copy the change accross to the object
' let the object validate the data input
'---------------------------------------------------------------------

   On Error GoTo ErrHandler
  If txtCategoryEdit.Visible Then
        Call UpdateCategoriesWithUserEdits
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "SelectedCRFFormId")
        Case OnErrorAction.Ignore
            Resume Next
        Case OnErrorAction.Retry
            Resume
        Case OnErrorAction.QuitMACRO
            Call ExitMACRO
            Call MACROEnd
    End Select
    
End Sub

'------------------------------------------
Private Sub txtCategoryEdit_LostFocus()
'------------------------------------------
' They've stopped editing in the text box
' Must update the grid to reflect the latest changes
'------------------------------------------

    ' Removed If mbRefreshCategoryGrid because this was
    ' preventing correct grid updates - NCJ 28/10/99
'    If mbRefreshCategoryGrid Then
        Call RefreshCategoriesGrid
'    End If
 
    
End Sub

'---------------------------------------------------------------------
Private Sub txtDataItemDerivation_Change()
'---------------------------------------------------------------------
    
    If Not blnDuringShow Then
        ' NCJ 6/3/01
        If txtDataItemDerivation.Tag <> txtDataItemDerivation.Text Then
            blnDataChanged = True
        End If
    End If

End Sub

'---------------------------------------------------------------------
Private Sub txtDataItemDerivation_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
' DPH 23/10/2001 - Throw away disallowed keystrokes
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    ' Disallow set of characters
    If InStr(1, msDisallowed, Chr(KeyAscii), vbBinaryCompare) > 0 Then
        KeyAscii = 0
    End If
    
    Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtDataItemDerivation_KeyPress")
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
Private Sub txtDataItemHelpText_Change()
'---------------------------------------------------------------------
    
    If Not blnDuringShow Then
        ' NCJ 6/3/01
        If txtDataItemHelpText.Tag <> txtDataItemHelpText.Text Then
            blnDataChanged = True
        End If
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub txtDataItemHelpText_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
' DPH 23/10/2001 - Throw away disallowed keystrokes
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    ' Disallow set of characters
    If InStr(1, msDisallowed, Chr(KeyAscii), vbBinaryCompare) > 0 Then
        KeyAscii = 0
    End If
    
    Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtDataItemHelpText_KeyPress")
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
Private Sub txtDataItemLength_Change()
'---------------------------------------------------------------------
' Only allow length to change if the format is empty
' NCJ 14/12/99, SR 2041 - Check for empty txtDataItemLength
'---------------------------------------------------------------------
Dim sLen As String
Dim nLen As Integer

    On Error GoTo ErrHandler

    If Not blnDuringShow Then
        
        If cboDataItemFormat <> vbNullString Then
            ' They've chosen a format so force the length on them
            nLen = Len(cboDataItemFormat)
            sLen = txtDataItemLength.Text
            If sLen > "" Then
                ' Change it back if it's different
                If CInt(sLen) <> nLen Then
                    txtDataItemLength.Text = nLen
                End If
            Else    ' Length is empty
                txtDataItemLength.Text = nLen
            End If
        End If
        
        ' NCJ 6/3/01
        If txtDataItemLength.Tag <> txtDataItemLength.Text Then
            blnDataChanged = True
        End If
        
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtDataItemLength_Change")
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
Private Sub txtDataItemLength_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler
    
    ' Allow only number 0 to 9
    ' And Backspace key (NCJ 18/11/99)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
    
    Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtDataItemLength_KeyPress")
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
Private Sub txtDataItemName_Change()
'---------------------------------------------------------------------
        
    If Not blnDuringShow Then
        ' NCJ 6/3/01
        If txtDataItemName.Tag <> txtDataItemName.Text Then
            blnDataChanged = True
        End If
    End If

End Sub

'---------------------------------------------------------------------
Private Sub txtDataItemName_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
' DPH 23/10/2001 - Throw away disallowed keystrokes
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    ' Disallow set of characters
    If InStr(1, msDisallowed, Chr(KeyAscii), vbBinaryCompare) > 0 Then
        KeyAscii = 0
    End If
    
    Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtDataItemName_KeyPress")
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
Private Sub txtDescription_Change()
'---------------------------------------------------------------------

    If blnDuringShow Then
        If txtDescription.Tag <> txtDescription.Text Then
            blnDataChanged = True
        End If
    End If
 
End Sub

'---------------------------------------------------------------------
Private Sub txtDescription_GotFocus()
'---------------------------------------------------------------------

    If Not mbForcedReFocus Then
        'Not a re-focus so set tag
        txtDescription.Tag = txtDescription.Text
    Else
        'It was a re-focus. Tag not changed and flag re-set
        mbForcedReFocus = False
    End If
End Sub

'---------------------------------------------------------------------
Private Sub txtDescription_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
    
    ' Disallow set of characters
    If InStr(1, msDisallowed, Chr(KeyAscii), vbBinaryCompare) > 0 Then
        KeyAscii = 0
    End If
End Sub

'---------------------------------------------------------------------
Private Sub txtDescription_LostFocus()
'---------------------------------------------------------------------

     If IsActiveControlCancel Then
        Exit Sub
    End If
    
    If Not gblnValidString(txtDescription.Text, valOnlySingleQuotes) Then
        MsgBox "Question description" & gsCANNOT_CONTAIN_INVALID_CHARS, _
                vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
        mbForcedReFocus = True
        txtDescription.SetFocus
        Exit Sub
    End If
    
    If txtDescription.Tag <> txtDescription.Text Then
        blnDataChanged = True
    End If
End Sub

'---------------------------------------------------------------------
Private Sub txtExportName_Change()
'---------------------------------------------------------------------

' NCJ 6 Mar 01 - Validation gets done in the LostFocus
'Dim sExportName As String
'    'WillC SR3741 17/8/00
'    sExportName = txtExportName.Text
'    If InStr(1, sExportName, " ", vbTextCompare) Then
'        MsgBox "You may not have spaces in the Export code.", vbInformation, "MACRO"
'        sExportName = Replace(sExportName, " ", "")
'        txtExportName.Text = sExportName
'    End If
    
    If Not blnDuringShow Then
        ' NCJ 6/3/01
        If txtExportName.Tag <> txtExportName.Text Then
            blnDataChanged = True
        End If
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub txtExportName_GotFocus()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    If Not mbForcedReFocus Then
        'Not a re-focus so set tag
        txtExportName.Tag = txtExportName.Text
    Else
        'It was a re-focus. Tag not changed and flag re-set
        mbForcedReFocus = False
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtExportName_GotFocus")
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
Private Sub txtExportName_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
' DPH 23/10/2001 - Throw away disallowed keystokes (here allow only ALPHANUMERIC)
' NCJ 16 Jun 03 - Rewritten for greater clarity, and to allow underscores (Bug 1355)
' NCJ 24 Jun 03 - Also allow Ctrl C, V and X
'---------------------------------------------------------------------
Const nCtrlC = 3
Const nCtrlV = 22
Const nCtrlX = 24
Const nCtrlZ = 26

    On Error GoTo ErrHandler
    
    ' Allow numbers 0 to 9
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    
    ' Allow letters a-z
    If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then Exit Sub
    
    ' Allow letters A-Z
    If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    
    ' Allow underscore
    If KeyAscii = Asc("_") Then Exit Sub
    
    ' Allow backspace
    If KeyAscii = vbKeyBack Then Exit Sub
    
    ' Allow delete
    If KeyAscii = vbKeyDelete Then Exit Sub
    
    ' NCJ 24 Jun 03 - Allow editing functions
    If KeyAscii = nCtrlC Then Exit Sub
    If KeyAscii = nCtrlV Then Exit Sub
    If KeyAscii = nCtrlX Then Exit Sub
    If KeyAscii = nCtrlZ Then Exit Sub
    
    ' Ignore all others
    KeyAscii = 0
    
'    ' And Backspace key (NCJ 18/11/99)
'    If Not ((KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 8) Or _
'     (KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii > 96 And KeyAscii < 123)) Then
'        KeyAscii = 0
'    End If
    
    Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtExportName_KeyPress")
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
Private Sub txtvalidationgridEdits_Change()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If txtValidationGridEdits.Visible Then
        Call UpdateValidationsWithUserEdits
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtvalidationgridEdits_Change")
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
Private Sub txtCategoryEdit_KeyDown(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    With gridValidation
        Select Case KeyCode
        Case vbKeyReturn
            '.SetFocus
            DoEvents
            If mbRefreshCategoryGrid Then
                Call RefreshCategoriesGrid
            End If
            txtAfterCategoryTextTabOrder.SetFocus
            
         End Select
     
     End With
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtCategoryEdit_KeyDown")
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
Private Sub txtValidationGridEdits_KeyDown(KeyCode As Integer, Shift As Integer)
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
   
  With gridValidation
    Select Case KeyCode
        Case vbKeyReturn
            DoEvents
            If mbRefreshValidationGrid Then
                Call RefreshValidationsGrid
            End If
            txtAfterValidationTextTabOrder.SetFocus
            
         End Select
     End With
     
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtValidationGridEdits_KeyDown")
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
Private Sub txtCategoryEdit_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
' DPH 23/10/2001 - Throw away disallowed keystokes
'---------------------------------------------------------------------
    On Error GoTo ErrHandler
   
    'noise suppression
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
    
    ' Disallow set of characters
    If gridCategories.Col = mcValueCodeCol Then
        ' If not Alphanumeric or underscore
        If Not ((KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 8) Or _
         (KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii > 96 And KeyAscii < 123) Or _
         (KeyAscii = 95)) Then
            KeyAscii = 0
        End If
    ElseIf gridCategories.Col = mcValueCol Then
        If InStr(1, msDisallowed, Chr(KeyAscii), vbBinaryCompare) > 0 Then
            KeyAscii = 0
        End If
    End If

    
    Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtCategoryEdit_KeyPress")
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
Private Sub EnableODBCControls(bEnable As Boolean)
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    lblConnect.Enabled = Not bEnable
    lblSQL.Enabled = Not bEnable
    txtSQL.Enabled = Not bEnable
    txtConnect.Enabled = Not bEnable
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "EnableODBCControls")
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
Private Sub sstDataDefinition_Click(PreviousTab As Integer)
'---------------------------------------------------------------------
'   Force DBGrid to accept focus and select column, so that delete works.
'---------------------------------------------------------------------

   On Error GoTo ErrHandler
  Select Case PreviousTab
    Case 1, 2
        Call HideEditingControls
    Case Else
        ' do nothing
    End Select
    
    If sstDataDefinition.Tab = 3 Then
        Call PopulateRequiredTab
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "sstDataDefinition_Click")
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
Private Sub EnableLoad()
'---------------------------------------------------------------------
' Enable or disable the Load button
' NCJ 13/1/00 - Check for trial status too
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
 
    ' NCJ 13/1/00 - Only enable for trail in preparation or new question
    If frmMenu.TrialStatus = eTrialStatus.InPreparation Or mbIsNew Then
        cmdLoad.Enabled = True
        If lstConnectionType.Text = "ADO - Default connection" Then
            If Trim$(txtConnect) = vbNullString Or Trim$(txtSQL) = vbNullString Then
                cmdLoad.Enabled = False
            End If
        End If
        
    Else
        cmdLoad.Enabled = False
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "EnableLoad")
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
Private Sub txtConnect_Change()
'---------------------------------------------------------------------

   Call EnableLoad
    
End Sub

'---------------------------------------------------------------------
Private Sub txtDataItemDerivation_GotFocus()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If Not mbForcedReFocus Then
        'Not a re-focus so set tag
        txtDataItemDerivation.Tag = txtDataItemDerivation.Text
    Else
        'It was a re-focus. Tag not changed and flag re-set
        mbForcedReFocus = False
    End If

Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtDataItemDerivation_GotFocus")
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
Private Sub txtDataItemDerivation_LostFocus()
'---------------------------------------------------------------------
' NCJ 17 Dec 99, SR 1768 - Do not validate string if user is clicking Cancel
'---------------------------------------------------------------------
Dim sDeriv As String

    On Error GoTo ErrHandler
    
    ' NCJ 17 Dec 99
    If IsActiveControlCancel Then
        Exit Sub
    End If

    If txtDataItemDerivation.Text = gsEMPTY_STRING Then
        ' Do nothing
    
    ElseIf Not (Me.ActiveControl.Name = "cmdDerivationExpression" _
         Or frmMenu.ActiveForm.Name = "frmCUIFunctionEditor") Then
    
        sDeriv = Trim(txtDataItemDerivation.Text)
        
        If Not gblnValidString(sDeriv, valOnlySingleQuotes) Then
            DialogWarning "Question derivation" & gsCANNOT_CONTAIN_INVALID_CHARS
            mbForcedReFocus = True
            txtDataItemDerivation.SetFocus
            Exit Sub
        End If
        
        If Len(sDeriv) > glMAX_AREZZO_EXPR_LEN Then
            DialogWarning "Question derivation may not be more than " & glMAX_AREZZO_EXPR_LEN & " characters"
            mbForcedReFocus = True
            txtDataItemDerivation.SetFocus
            Exit Sub
        End If
        
        If Not gclmGuideline.IsValidExpression(sDeriv) Then
            DialogWarning "Question derivation is not a valid expression."
            mbForcedReFocus = True
            txtDataItemDerivation.SetFocus
            Exit Sub
        End If

    End If

    If txtDataItemDerivation.Tag <> txtDataItemDerivation.Text Then
        blnDataChanged = True
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtDataItemDerivation_LostFocus")
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
Private Sub txtDataItemHelpText_GotFocus()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    If Not mbForcedReFocus Then
        'Not a re-focus so set tag
        txtDataItemHelpText.Tag = txtDataItemHelpText.Text
    Else
        'It was a re-focus. Tag not changed and flag re-set
        mbForcedReFocus = False
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtDataItemHelpText_GotFocus")
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
Private Sub txtDataItemHelpText_LostFocus()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    ' NCJ 6 Mar 01 - Don't validate if user is clicking Cancel
    If IsActiveControlCancel Then
        Exit Sub
    End If
    
    If Not gblnValidString(txtDataItemHelpText.Text, valOnlySingleQuotes) Then
        MsgBox "Question help text" & gsCANNOT_CONTAIN_INVALID_CHARS, _
                vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
        mbForcedReFocus = True
        txtDataItemHelpText.SetFocus
        Exit Sub
    End If
    
    If txtDataItemHelpText.Tag <> txtDataItemHelpText.Text Then
        blnDataChanged = True
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtDataItemHelpText_LostFocus")
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
Private Sub txtDataItemLength_GotFocus()
'---------------------------------------------------------------------
 On Error GoTo ErrHandler

    If Not mbForcedReFocus Then
        'Not a re-focus so set tag
        txtDataItemLength.Tag = txtDataItemLength.Text
    Else
        'It was a re-focus. Tag not changed and flag re-set
        mbForcedReFocus = False
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtDataItemLength_GotFocus")
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
Private Sub txtDataItemLength_LostFocus()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler

    If Not blnDuringShow Then
        If txtDataItemLength.Tag <> txtDataItemLength.Text Then blnDataChanged = True
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtDataItemLength_LostFocus")
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
Public Sub CloseWindow()
'---------------------------------------------------------------------
   
   Call HideWindow
        
End Sub

'---------------------------------------------------------------------
Private Sub cmdCancel_Click()
'---------------------------------------------------------------------

Dim oForm As Form

    mbCancelled = True
    If mbIsNew Then
        '   NCJ 15 Jan 02 - Ensure we always use correct instance of frmDataList (from old 3.0)

        Set oForm = frmMenu.gFindQuestionList(Me.ClinicalTrialId, Me.VersionId)
        If Not oForm Is Nothing Then
            ' Delete the question from the data list
            Call oForm.DeleteDataItem
        End If
    End If
   
   Call HideWindow
   mbCancelled = False
   
        
End Sub

'---------------------------------------------------------------------
Public Sub HideWindow()
'---------------------------------------------------------------------
' NCJ 22/12/99 - Rebuild the data item on current form if it's visible
'---------------------------------------------------------------------
Dim oForm As Form

    On Error GoTo ErrHandler
    
    Call HideEditingControls
    frmDataDefinition.Hide
    DoEvents
    blnNameChanged = False
    blnDataChanged = False
    Call DestroyCollections
    mbIsNew = False
    ' NCJ Refresh question if frmCRFDesign is already loaded
    ' (Don't do it if not because we get the wrong CRFs! SR 2568)
    Set oForm = frmMenu.gFindForm(Me.ClinicalTrialId, Me.VersionId, "frmCRFDesign")
    If Not oForm Is Nothing Then
        Call frmCRFDesign.RebuildQuestionIfVisible(Me.DataItemId)
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "HideWindow")
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
Private Sub cmdDerivationExpression_Click()
'---------------------------------------------------------------------
' Display the AREZZO function editor
' NCJ 28 Jan 03 - Say what data type we're expecting
' ic 13/06/2005 added clinical coding
'---------------------------------------------------------------------
Dim sArezzoType As String

    On Error GoTo ErrHandler

    Select Case menDataType
    'ic 13/06/2005 clinical coding: added thesaurus datatype
    Case DataType.Text, DataType.Thesaurus
        sArezzoType = "string"
    Case DataType.IntegerData, DataType.Real, DataType.LabTest
        sArezzoType = "numeric"
    Case DataType.Date
        sArezzoType = "temporal"
    End Select
    
    frmCUIFunctionEditor.Initialise txtDataItemDerivation, "Expression", _
        "Derivation expression for question " & Me.DataItemName, sArezzoType
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "cmdDerivationExpression_Click")
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

    If KeyCode = vbKeyF1 Then               ' Show user guide
        'ShowDocument Me.hWnd, gsMACROUserGuidePath
        
        'REM 07/12/01 - New Call for MACRO Help
        Call MACROHelp(Me.hWnd, App.Title)
    End If
    
End Sub

'---------------------------------------------------------------------
Private Sub txtDataItemName_GotFocus()
'---------------------------------------------------------------------
    
    On Error GoTo ErrHandler

    If Not mbForcedReFocus Then
        'Not a re-focus so set tag
        txtDataItemName.Tag = txtDataItemName.Text
    Else
        'It was a re-focus. Tag not changed and flag re-set
        mbForcedReFocus = False
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtDataItemName_GotFocus")
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
Private Sub txtDataItemName_LostFocus()
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    ' NCJ 6 Mar 01 - Don't validate if user is clicking Cancel
    If IsActiveControlCancel Then
        Exit Sub
    End If
    
    'do not allow an empty string or a leading spaces. JL 3/9/98 .Bug no 425.
    If Trim(txtDataItemName.Text) = "" Then
        MsgBox "Please enter a question name"
        mbForcedReFocus = True
        txtDataItemName.SetFocus
        Exit Sub
    End If
    
    If Not gblnValidString(txtDataItemName.Text, valOnlySingleQuotes) Then
        ' NCJ 31/5/00 - Added tildes to message
        MsgBox "A question name" & gsCANNOT_CONTAIN_INVALID_CHARS, _
                vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
        mbForcedReFocus = True
        txtDataItemName.SetFocus
        Exit Sub
    End If
    
    If txtDataItemName.Tag <> txtDataItemName.Text Then
        blnNameChanged = True
        blnDataChanged = True
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtDataItemName_LostFocus")
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
Private Function GetCaseOptions() As Integer
'---------------------------------------------------------------------
' Get the "Case" setting
' 0 = Proper (or undefined), 1 = Upper, 2 = Lower
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
   If optCase(0).Value Then
        GetCaseOptions = 0
    ElseIf optCase(1).Value Then
        GetCaseOptions = 1
    ElseIf optCase(2).Value Then
        GetCaseOptions = 2
    Else
        GetCaseOptions = 0
    End If
    
Exit Function
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "GetCaseOptions")
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
Private Sub SetCaseTestOptions(nDataType As DataType, Optional lCase As Long = 0, Optional sClinicalTestCode As String = "", _
    Optional nDictionarycode As Integer = -1)
'---------------------------------------------------------------------
' If nDatatype is text - set "Case" options
' lCase = 0 means unspecified case
' lCase = 1 means Upper case
' lCase = 2 means Lower case
' if Datatype is LabTest then sClinicalTestCode Clinical Test Code
' else disable lab and case frames
' ic 15/06/2005 added clinical coding
'---------------------------------------------------------------------
Dim n As Integer

    On Error GoTo ErrHandler
    
    Select Case nDataType
    Case DataType.Text
        fraCoding.Visible = False
        fraCase.Visible = True
        fraTest.Visible = False
        optCase(lCase).Value = True
    Case DataType.LabTest
        fraCoding.Visible = False
        fraCase.Visible = False
        fraTest.Visible = True
        'stuff to set lab
        If sClinicalTestCode = "" And cboGroup.ListCount > 0 Then
            'set to first group
             cboGroup.ListIndex = 0
        Else
            cboGroup.Text = moClinicalTests.Item(sClinicalTestCode).ClinicalTestGroupCode
            cboTest.Text = sClinicalTestCode
        End If
        cboTest_Click
    Case DataType.Thesaurus
        'ic 15/06/2005 clinical coding: added thesaurus handler
        fraCase.Visible = False
        fraTest.Visible = False
        fraCoding.Visible = True
        If (cboDictionary.ListCount = 0) Then
            'no dictionarys installed - disable the dictionary frame
            fraCoding.Enabled = False
            cboDictionary.Enabled = False
            
        Else
            'dictionaries installed - enable dictionary frame and select appropriate initial values
            fraCoding.Enabled = True
            cboDictionary.Enabled = True
            
            cboDictionary.ListIndex = 0
            If (nDictionarycode > -1) Then
                'search for the entry with the itemdata that matches this dictionary code
                For n = 0 To cboDictionary.ListCount - 1
                    If cboDictionary.ItemData(n) = nDictionarycode Then
                        cboDictionary.ListIndex = n
                        Exit For
                    End If
                Next n
            End If
        End If
        
    Case Else
        fraCoding.Visible = False
        fraCase.Visible = False
        fraTest.Visible = False
        
    End Select
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, "SetCaseTestOptions")
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
Public Sub UpdateDataDefinition()
'---------------------------------------------------------------------
' Save the current values to the MACRO database
' Assume everything is valid
' NCJ 17 Nov 99 - Convert format string to "standard" before saving
' ic 15/06/2005 added clinical coding
'---------------------------------------------------------------------
Dim nDataItemLength As Integer
Dim sCase As String

Dim sDataItemCode As String
Dim sDataItemName As String
Dim nDataItemType As Integer
Dim sDerivation As String
Dim sUnit As String
Dim nTrialTypeId As Integer
Dim nRequired As Integer
Dim bLibraryMode As Boolean
Dim sFormat As String
Dim sClinicalTestCode As String
Dim oForm As Form
Dim enMACROOnly As eMACROOnly
Dim sDescription As String
Dim nDictionarycode As Integer


    On Error GoTo ErrHandler

    
    ' PN change 41
    txtDataItemName.SetFocus
    sCase = GetCaseOptions()
    
    ' NCJ 12/8/99
    sDataItemCode = txtDataItemCode.Text
    sDataItemName = txtDataItemName.Text
    nDataItemType = menDataType
    sDerivation = txtDataItemDerivation.Text
    sUnit = cboUnitOfMeasurement.Text
    
    
    'ZA 19/08/2002 - set MACROOnly and Description properties
    If chkMACROOnly.Value = vbChecked Then
        enMACROOnly = eMACROOnly.MACROOnly
    Else
        enMACROOnly = eMACROOnly.NotMACROOnly
    End If
    sDescription = txtDescription.Text
    
    ' Set the data item length according to data type NCJ 28/10/99
    Select Case nDataItemType
    Case DataType.Category
        nDataItemLength = moCategories.MaxCatLength
    ' NCJ 9 Feb 00 - Set date fields to predefined lengths
    ' according to type of date format
    ' mnDateFormatType should have been set up in ValidateDataFormat
    ' This is in case they've specified e.g. "d/m/yyyy" which actually needs length of 10
    Case DataType.Date
        Select Case mnDateFormatType
        Case eDateFormatType.DateOnly
            nDataItemLength = Len("dd/mm/yyyy")     ' Longest date format
        Case eDateFormatType.TimeOnly
            nDataItemLength = Len("hh:mm:ss")     ' Longest time format
        Case eDateFormatType.DateAndTime
            nDataItemLength = Len("dd/mm/yyyy hh:mm:ss")     ' Longest datetime format
        Case Else
            ' Shouldn't get here!
            nDataItemLength = Len("dd/mm/yyyy hh:mm:ss")
        End Select
    Case Else
        If txtDataItemLength.Text > "" Then
            nDataItemLength = txtDataItemLength.Text
        ElseIf cboDataItemFormat.Text > "" Then
            nDataItemLength = Len(cboDataItemFormat.Text)
        Else
            nDataItemLength = 0
        End If
'        If txtDataItemLength.Text = "" Then
'            nDataItemLength = 0
'        Else
'            nDataItemLength = txtDataItemLength.Text
'        End If
    End Select

    ' PN 17/09/99
    ' save the trial type
    bLibraryMode = frmMenu.IsAppInLibraryMode
    If bLibraryMode Then
        Call SaveRequiredTrials
'
'        nTrialTypeId = moTrialTypes.Key(cboTrialTypes)
'        nRequired = 1
    
    End If
    
    ' NCJ 17 Nov 99 - Convert numeric format string before saving
    ' NCJ 25/9/00 - Include LabTest type
    Select Case nDataItemType
    Case DataType.IntegerData, DataType.Real, DataType.LabTest
        ' NCJ 18 Jun 02 - CBB 2.2.14/19 Added extra parameter TRUE
        ' to preserve the thousand separators
        sFormat = ConvertLocalNumToStandard(cboDataItemFormat.Text, True)
    Case Else
        sFormat = cboDataItemFormat.Text
    End Select
    
    'TA 25/09/200: only get clinical test id if it's a lab test question
    If nDataItemType = DataType.LabTest Then
        'get test id
        sClinicalTestCode = cboTest.Text
    Else
        'value of "" means no test id to be set
        sClinicalTestCode = ""
    End If
    
    'ic 15/06/2005 clinical coding: get chosen dictionary
    If (nDataItemType = DataType.Thesaurus) Then
        nDictionarycode = cboDictionary.ItemData(cboDictionary.ListIndex)
    Else
        nDictionarycode = -1
    End If
    
    'ic 15/06/2005 clinical coding: added dictionary choice
    Call gdsUpdateDataItem(ClinicalTrialId, _
        VersionId, mlDataItemId, sDataItemCode, sDataItemName, _
        nDataItemType, nDataItemLength, sFormat, sUnit, _
        sDerivation, txtDataItemHelpText.Text, sCase, txtExportName, _
        nRequired, nTrialTypeId, bLibraryMode, sClinicalTestCode, enMACROOnly, sDescription, nDictionarycode)
    
    ' Added NCJ 12/8/99
    UpdateProformaDataItem mlDataItemId, sDataItemCode, _
                        sDataItemName, nDataItemType, sDerivation, sUnit
      
    If blnNameChanged _
    Or txtDataItemName.Text <> txtDataItemName.Tag Then
        Call frmMenu.ViewQuestions(gsUPDATE, frmMenu.ClinicalTrialId, frmMenu.ClinicalTrialName) 'SDM SR2063 01/11/99   Ensure that the form is available
        ' NCJ 15 Jan 02 - Make sure ALL question lists are updated
        For Each oForm In Forms
            If oForm.Name = "frmDataList" Then
                If oForm.ClinicalTrialId = Me.ClinicalTrialId And oForm.VersionId = Me.VersionId Then
                    Call oForm.ChangeDataItemName(mlDataItemId, txtDataItemName, txtDataItemCode)
                End If
            End If
        Next
    End If
    
    UpdateControlType ClinicalTrialId, VersionId, mlDataItemId, menDataType

' WillC 25/2/00 SR2990 Removed the rebuilding the entire page
'  instead of just the question that was being changed.
    
    ' save the category data
    moCategories.Save
    moValidations.Save
    mbIsNew = False
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "UpdateDataDefinition")
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
Private Sub UpdateControlType(ByVal vClinicalTrialId As Long, _
                             ByVal vVersionId As Integer, _
                             ByVal vDataItemId As Long, _
                             ByVal vDataItemType As Integer)
'---------------------------------------------------------------------
' NCJ 25/9/00 - Changed from Public to Private because only used in this module
' ic 13/06/2005 added clinical coding
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsCRFElement As ADODB.Recordset
Dim nControlType As Integer

    On Error GoTo ErrHandler
    
    sSQL = "SELECT ControlType FROM CRFELEMENT " _
        & " WHERE ClinicalTrialId = " & vClinicalTrialId _
        & "   AND VersionId = " & vVersionId _
        & "   AND DataItemId = " & vDataItemId
        
    Set rsCRFElement = New ADODB.Recordset
    rsCRFElement.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockPessimistic, adCmdText
    ' NB We only want to reset the control type if the user has't already set it
    'SDM SR2059 03/11/99 Added gn_TEXT_BOX and gnPUSH_BUTTONS + gnOPTION_BUTTONS to If
    ' NCJ 7 Apr 03 MACRO 3.0 Bug 1439 - Changed control types and added Text Box back in
    Do While Not rsCRFElement.EOF
        nControlType = rsCRFElement!ControlType
            Select Case vDataItemType
            Case DataType.Category
                If nControlType <> gn_POPUP_LIST _
                And nControlType <> gn_PUSH_BUTTONS _
                And nControlType <> gn_OPTION_BUTTONS _
                And nControlType <> gn_TEXT_BOX Then
                    ' Default to popup list
                    rsCRFElement!ControlType = gnPOPUP_LIST
                    rsCRFElement.Update
                End If
            Case DataType.Date
                If nControlType <> gn_TEXT_BOX _
                And nControlType <> gn_CALENDAR Then
                    ' NCJ 29/12/99, SR 2423 - Set default to be gn_TEXT_BOX
'                    rsCRFElement!ControlType = gn_CALENDAR
                    rsCRFElement!ControlType = gn_TEXT_BOX
                    rsCRFElement.Update
                End If
                ' NCJ 25/9/00 - Include LabTest here
                'ic 13/06/2005 clinical coding: added thesaurus datatype
            Case DataType.IntegerData, DataType.Real, DataType.Text, DataType.LabTest, DataType.Thesaurus
                If nControlType <> gn_TEXT_BOX Then
                    rsCRFElement!ControlType = gn_TEXT_BOX
                    rsCRFElement.Update
                End If
            Case DataType.Multimedia
                If nControlType <> gn_ATTACHMENT Then
                    rsCRFElement!ControlType = gn_ATTACHMENT
                    rsCRFElement.Update
                End If
            End Select
    
        rsCRFElement.MoveNext
    Loop
    
    rsCRFElement.Close
    Set rsCRFElement = Nothing
    
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmDataDefinition.UpdateControlType"
    
End Sub

'---------------------------------------------------------------------
Private Function ChangesValid() As Boolean
'---------------------------------------------------------------------
' this function returns true if the validation and category grid edits
' can be saved
'---------------------------------------------------------------------
Dim bChangesValid As Boolean
    
    On Error GoTo ErrHandler
    
    bChangesValid = True

    ' check if category grid edits can be saved
    If Not moCategories Is Nothing Then
        bChangesValid = moCategories.IsValid
    End If
    
    If bChangesValid Then
        ' check if validation grid edits can be saved
        If Not moValidations Is Nothing Then
            bChangesValid = moValidations.IsValid
        End If
    End If
    
    If bChangesValid Then
        bChangesValid = (IsDefinitionValid = AllValid)
    End If
    
    ChangesValid = bChangesValid
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmDataDefinition.ChangesValid"
    
End Function

'---------------------------------------------------------------------
Private Function HasChanges() As Boolean
'---------------------------------------------------------------------
' determine if the data has changes applied
'---------------------------------------------------------------------

    On Error GoTo ErrHandler
    
    If Not blnDataChanged Then
        ' controls have not changed
        
        If Not moCategories Is Nothing Then
            ' have the categories changed
            blnDataChanged = moCategories.IsDirty
        
        End If
    
        If Not moValidations Is Nothing Then
            If Not blnDataChanged Then
                ' have the validations changed
                blnDataChanged = moValidations.IsDirty
            End If
            
        End If
    End If
    HasChanges = blnDataChanged
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmDataDefinition.HasChanges"
    
End Function

'---------------------------------------------------------------------
Public Function SaveChangedData(Optional bDoSave As Boolean = True) As Boolean
'---------------------------------------------------------------------
' this function handles the logic of what to do when the user opts to
' display a different data item or close the window
'---------------------------------------------------------------------
Dim sMsg As String
Dim iMsgboxButtons As Integer

    On Error GoTo ErrHandler

    If HasChanges Then
        ' changes have been made to the form controls or grids
        If ChangesValid Then
            If mbIsNew Then
                ' new items must be saved or cancel action
                ' RS 13/02/2003 Changed Me.DataItemName to txtDataItemCode.Text
                sMsg = "This is a new question definition '" & txtDataItemCode.Text
                sMsg = sMsg & "' and must be saved now to continue. Do you wish to continue?"
                iMsgboxButtons = vbOKCancel
                
            Else
                ' existing items do not have to be saved
                ' ASH 21/06/2002 Changed Me.DataItemName to txtDataItemCode.Text SR4640
                sMsg = "Do you want to save question definition '" & txtDataItemCode.Text & "'  ?"
                iMsgboxButtons = vbYesNoCancel

            End If
            
            Select Case MsgBox(sMsg, iMsgboxButtons + vbApplicationModal + vbQuestion, gsDIALOG_TITLE)
            Case vbYes, vbOK
                If bDoSave Then
                    UpdateDataDefinition
                    blnDataChanged = False
                
                End If
                mbDiscardChanges = True
                
            Case vbNo
                ' no need to save
                ' allow the caller to complete their action
                mbDiscardChanges = True
                
            Case vbCancel
                ' notify the caller to cancel the action
                mbDiscardChanges = False
            
            End Select
                
        Else
            ' the changes are not valid
            If mbIsNew Then
                Call ShowInvalidDefinitionMessage
                
                ' notify the caller to cancel the action
                mbDiscardChanges = False

            Else
                ' ask the user if the changes should be discarded
                ' RS 13/02/2003 Changed Me.DataItemName to txtDataItemCode.Text
                sMsg = "The changes to question definition '" & txtDataItemCode.Text & "' are" & vbCr
                sMsg = sMsg & "not valid. If you continue you will lose these changes."
                Select Case MsgBox(sMsg, vbOKCancel + vbApplicationModal + vbQuestion, gsDIALOG_TITLE)
                Case vbOK
                    ' no need to save
                    ' allow the caller to complete their action
                    mbDiscardChanges = True

                Case vbCancel
                    ' notify the caller to cancel the action
                    mbDiscardChanges = False
                    
                End Select
            
            End If
        
        End If
        
    Else
        ' No changes have been made
        ' allow the caller to complete their action
        mbDiscardChanges = True
        
    End If
    
    SaveChangedData = mbDiscardChanges
    
Exit Function
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmDataDefinition.SaveChangedData"
    
End Function

'---------------------------------------------------------------------
Private Sub ShowInvalidDefinitionMessage()
'---------------------------------------------------------------------
   Dim sMsg As String

    ' and this is a new item
    ' so the save must be done to save a valid max length
    sMsg = "Question definition '" & Me.DataItemName & "' must be valid before" & vbCr
    sMsg = sMsg & "any other question definitions can be displayed."
    Call DialogInformation(sMsg)
    
End Sub

'---------------------------------------------------------------------
Private Sub txtExportName_LostFocus()
'---------------------------------------------------------------------
Dim sMsg As String
   
    On Error GoTo ErrHandler
    
    ' NCJ 6 Mar 01 - Don't validate if user is clicking Cancel
    If IsActiveControlCancel Then
        Exit Sub
    End If
    
    sMsg = ValidateDataItemExportName(txtExportName.Text, mlClinicalTrialId, _
                                    mnVersionId, txtDataItemCode.Text)
    If sMsg <> vbNullString Then
        Call DialogError(sMsg)
        mbForcedReFocus = True
        txtExportName.SetFocus
    End If
    
    If txtExportName.Tag <> txtExportName.Text Then
        blnDataChanged = True
    End If
    
Exit Sub
ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtExportName_LostFocus")
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
Private Sub txtSQL_Change()
'---------------------------------------------------------------------

    Call EnableLoad

End Sub

'---------------------------------------------------------------------
Private Sub txtValidationGridEdits_KeyPress(KeyAscii As Integer)
'---------------------------------------------------------------------
' DPH 23/10/2001 - Throw away disallowed keystokes
'---------------------------------------------------------------------
    On Error GoTo ErrHandler
   
    'noise suppression
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
    
    If InStr(1, msDisallowed, Chr(KeyAscii), vbBinaryCompare) > 0 Then
        KeyAscii = 0
    End If
    
    Exit Sub

ErrHandler:
    Select Case MACROFormErrorHandler(Me, Err.Number, Err.Description, _
                                    "txtValidationGridEdits_KeyPress")
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
Private Sub txtValidationGridEdits_LostFocus()
'---------------------------------------------------------------------
' They've stopped editing a validation
'---------------------------------------------------------------------
Dim nFocusHandle As Long
    ' Removed If mbRefreshValidationGrid which was
    ' preventing correct grid updates - NCJ 28/10/99
'    If mbRefreshValidationGrid Then
        Call RefreshValidationsGrid
        ' DPH 16/10/2001 - Code added to hide the data entry control(s) to fix the resize display problem
        ' DPH 17/10/2001 - API call for focus
        ' DPH 23/10/2001 - More tweaking for Lost Focus Event
        'txtValidationGridEdits.Visible = False
        nFocusHandle = GetFocus
        If cmdValidationExpression.Visible And nFocusHandle <> cmdValidationExpression.hWnd Then
            cmdValidationExpression.Visible = False
        End If
        If nFocusHandle <> cmdValidationExpression.hWnd Then
            txtValidationGridEdits.Visible = False
        End If
'    End If
    
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
    'WillC SR3573   changed heading form Trial to Study
    lsvRequired.ColumnHeaders.Add , , "Study"
    sSQL = "SELECT * FROM TrialType"
    Set rsRequiredData = New ADODB.Recordset
    rsRequiredData.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsRequiredData.EOF Then
        Set rsRequiredData = Nothing
        Exit Sub
    End If
    rsRequiredData.MoveFirst
    Do While Not rsRequiredData.EOF
        lsvRequired.ListItems.Add , Trim(Str(rsRequiredData!TrialTypeId) & "K"), rsRequiredData!TrialTypeName
        rsRequiredData.MoveNext
    Loop
    
    'Set check marks
    sSQL = "SELECT TrialTypeId FROM RequiredData WHERE DataItemId = " & mlDataItemId
    Set rsRequiredData = New ADODB.Recordset
    rsRequiredData.Open sSQL, MacroADODBConnection, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rsRequiredData.EOF Then
        lsvRequired.SelectedItem.Selected = False
        Set rsRequiredData = Nothing
        Exit Sub
    End If
    rsRequiredData.MoveFirst
    Do While Not rsRequiredData.EOF
        lsvRequired.ListItems(rsRequiredData!TrialTypeId & "K").Checked = True
        rsRequiredData.MoveNext
    Loop
    Set rsRequiredData = Nothing
    
    lsvRequired.SelectedItem.Selected = False
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmDataDefinition.PopulateRequiredTab"

End Sub

'---------------------------------------------------------------------
Private Sub SaveRequiredTrials()
'---------------------------------------------------------------------
    'SDM 14/12/99 SR1177
'---------------------------------------------------------------------
Dim sSQL As String
Dim rsRequiredData As ADODB.Recordset
Dim liTrialType As MSComctlLib.ListItem
 
    On Error GoTo ErrHandler
    
    sSQL = "SELECT * FROM RequiredData WHERE DataItemId =  " & mlDataItemId
    Set rsRequiredData = New ADODB.Recordset
    rsRequiredData.Open sSQL, MacroADODBConnection, adOpenDynamic, adLockOptimistic, adCmdText
    lsvRequired.Visible = False
    lsvRequired.Sorted = False
    For Each liTrialType In lsvRequired.ListItems
        rsRequiredData.Filter = "DataItemId = " & mlDataItemId & " AND TrialTypeId = " & Val(liTrialType.Key)
        If liTrialType.Checked Then
            If rsRequiredData.EOF Then
                rsRequiredData.AddNew
            End If
            rsRequiredData!DataItemId = mlDataItemId
            rsRequiredData!TrialTypeId = Val(liTrialType.Key)
            rsRequiredData.Update
        Else
            If Not rsRequiredData.EOF Then
                rsRequiredData.Delete
            End If
        End If
    Next liTrialType
   lsvRequired.Visible = True
   lsvRequired.Sorted = True
Exit Sub
ErrHandler:
    Err.Raise Err.Number, , Err.Description & "|frmDataDefinition.SaveRequiredTrials"

End Sub

'---------------------------------------------------------------------
Private Function PromptDeleteGridRow(bCheckTrialStatus As Boolean) As VbMsgBoxResult
'---------------------------------------------------------------------
' PN 18/08/99 - routine added
' Prompt the user to delete a grid row
' NCJ 12/1/00 - Check trial status (code moved here from basCommon)
'---------------------------------------------------------------------
Dim sMsg As String
    
    If bCheckTrialStatus And frmMenu.TrialStatus > eTrialStatus.InPreparation _
     And Not mbIsNew Then
        sMsg = "You may not delete category values once a study has been opened." & vbCrLf
        sMsg = sMsg & "You may make category values inactive instead."
        MsgBox sMsg, vbOKOnly + vbExclamation + vbDefaultButton1 + vbApplicationModal, gsDIALOG_TITLE
        PromptDeleteGridRow = False
    Else
        sMsg = "This action will delete the selected row(s). Do you wish to continue?"
        PromptDeleteGridRow = DialogQuestion(sMsg)
    End If
    
End Function

'---------------------------------------------------------------------
Private Sub SetUnitsComboText(sUnit As String)
'---------------------------------------------------------------------
' set text of combo assuming the first item is an empty string
'---------------------------------------------------------------------

    If sUnit = "" Then
        cboUnitOfMeasurement.ListIndex = 0
    Else
        cboUnitOfMeasurement.Text = sUnit
    End If
End Sub

'---------------------------------------------------------------------
Private Function IsActiveControlCancel() As Boolean
'---------------------------------------------------------------------
' check if Active control is cmdCancel
' REM 23/07/02 - CBB 2.2.20 No. 10
'---------------------------------------------------------------------

    If Me.ActiveControl Is Nothing Then
        IsActiveControlCancel = False
    Else
        IsActiveControlCancel = (Me.ActiveControl.Name = "cmdCancel")
    End If
    
End Function

'---------------------------------------------------------------------
Private Function IsFormVisitDateQuestion(ByVal lTrialId As Long, ByVal nVersionId As Integer, _
                                          ByVal lDataItemId As Long) As Boolean
'---------------------------------------------------------------------
'ZA 21/08/2002 - checks if the question is being used as a form or visit date
'---------------------------------------------------------------------
Dim rsCRFElement As ADODB.Recordset
Dim sQuery As String
    
    IsFormVisitDateQuestion = False
    
    sQuery = "select ElementUse from CRFElement where ClinicalTrialId = " & lTrialId & " and VersionId = " & nVersionId & _
             " and DataItemId  = " & lDataItemId
    
    Set rsCRFElement = New ADODB.Recordset
    
    rsCRFElement.Open sQuery, MacroADODBConnection, adOpenKeyset, adLockOptimistic
    
    If rsCRFElement.EOF Or rsCRFElement.BOF Then
        Exit Function
    Else
        Do While Not rsCRFElement.EOF
            If rsCRFElement.Fields("ElementUse").Value = 1 Then
                IsFormVisitDateQuestion = True
                Exit Function
            End If
            rsCRFElement.MoveNext
        Loop
    End If
    rsCRFElement.Close
    Set rsCRFElement = Nothing
    
End Function

