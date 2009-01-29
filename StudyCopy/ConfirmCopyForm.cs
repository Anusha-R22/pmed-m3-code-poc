using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

//----------------------------------------------------------------------
// 21/09/2006	bug 2804 check for selected properties checkboxes
// 21/09/2006 bug 2806 clear copyproperties array
//----------------------------------------------------------------------

namespace InferMed.MACRO.StudyCopy
{
	/// <summary>
	/// Summary description for ConfirmCopyForm.
	/// </summary>
	public class ConfirmCopyForm : System.Windows.Forms.Form
	{
		//settings file
		private string _settingsFile = "";

		private System.Windows.Forms.GroupBox groEform;
		private System.Windows.Forms.Button btnNo;
		private System.Windows.Forms.Button btnYes;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.CheckBox chkEformBGColour;
		private System.Windows.Forms.CheckBox chkEformLabel;
		private System.Windows.Forms.CheckBox chkEformLocalLabel;
		private System.Windows.Forms.CheckBox chkEformDisplayNumbers;
		private System.Windows.Forms.CheckBox chkEformHideIfInactive;
		private System.Windows.Forms.CheckBox chkEformWidth;
		private System.Windows.Forms.GroupBox groElement;
		private System.Windows.Forms.CheckBox chkElementControlType;
		private System.Windows.Forms.CheckBox chkElementFontColour;
		private System.Windows.Forms.CheckBox chkElementCaption;
		private System.Windows.Forms.CheckBox chkElementFontName;
		private System.Windows.Forms.CheckBox chkElementFontBold;
		private System.Windows.Forms.CheckBox chkElementFontItalic;
		private System.Windows.Forms.CheckBox chkElementFontSize;
		private System.Windows.Forms.CheckBox chkElementFieldOrder;
		private System.Windows.Forms.CheckBox chkElementSkipCondition;
		private System.Windows.Forms.CheckBox chkElementHeight;
		private System.Windows.Forms.CheckBox chkElementWidth;
		private System.Windows.Forms.CheckBox chkElementCaptionX;
		private System.Windows.Forms.CheckBox chkElementCaptionY;
		private System.Windows.Forms.CheckBox chkElementX;
		private System.Windows.Forms.CheckBox chkElementY;
		private System.Windows.Forms.CheckBox chkElementPrintOrder;
		private System.Windows.Forms.CheckBox chkElementHidden;
		private System.Windows.Forms.CheckBox chkElementLocalFlag;
		private System.Windows.Forms.CheckBox chkElementOptional;
		private System.Windows.Forms.CheckBox chkElementMandatory;
		private System.Windows.Forms.CheckBox chkElementRequireComment;
		private System.Windows.Forms.CheckBox chkElementRoleCode;
		private System.Windows.Forms.CheckBox chkElementOwnerQGroupId;
		private System.Windows.Forms.CheckBox chkElementQGroupId;
		private System.Windows.Forms.CheckBox chkElementQGroupFieldOrder;
		private System.Windows.Forms.CheckBox chkElementShowStatusFlag;
		private System.Windows.Forms.CheckBox chkElementUse;
		private System.Windows.Forms.CheckBox chkElementDisplayLength;
		private System.Windows.Forms.CheckBox chkElementCaptionFontName;
		private System.Windows.Forms.CheckBox chkElementCaptionFontBold;
		private System.Windows.Forms.CheckBox chkElementCaptionFontItalic;
		private System.Windows.Forms.CheckBox chkElementCaptionFontSize;
		private System.Windows.Forms.CheckBox chkElementCaptionFontColour;
		private System.Windows.Forms.CheckBox chkEformTitle;
		private System.Windows.Forms.GroupBox groDataItem;
		private System.Windows.Forms.CheckBox chkDataItemValidation;
		private System.Windows.Forms.CheckBox chkDataItemDescription;
		private System.Windows.Forms.CheckBox chkDataItemCase;
		private System.Windows.Forms.CheckBox chkDataItemHelpText;
		private System.Windows.Forms.CheckBox chkDataItemDerivation;
		private System.Windows.Forms.CheckBox chkDataItemLength;
		private System.Windows.Forms.CheckBox chkDataItemUnitOfMeasurement;
		private System.Windows.Forms.CheckBox chkDataItemFormat;
		private System.Windows.Forms.CheckBox chkDataItemDataType;
		private System.Windows.Forms.GroupBox groQG;
		private System.Windows.Forms.CheckBox chkQGMaxRepeats;
		private System.Windows.Forms.CheckBox chkQGMinRepeats;
		private System.Windows.Forms.CheckBox chkQGInitialRows;
		private System.Windows.Forms.CheckBox chkQGDisplayRows;
		private System.Windows.Forms.CheckBox chkQGBorder;
		private System.Windows.Forms.CheckBox chkQGDisplayType;
		private System.Windows.Forms.CheckBox chkQGName;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ConfirmCopyForm( string settingsFile )
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//settings file
			_settingsFile = settingsFile;
		}

		public DialogResult ShowDialog( StudyCopyGlobal.ElementType eType )
		{
			_eType = eType;
			this.DialogResult = DialogResult.No;
			_elementCopyProperties.Clear();
			_dataItemCopyProperties.Clear();
			_qgCopyProperties.Clear();
			_eformQGCopyProperties.Clear();

			switch( _eType )
			{
				case StudyCopyGlobal.ElementType.eForm:
					groEform.Visible = true;
					groElement.Visible = false;
					groDataItem.Visible = false;
					groQG.Visible = false;
					this.Height = 200;
					break;
				case StudyCopyGlobal.ElementType.eFormElement:
					groEform.Visible = false;
					groElement.Visible = true;
					groDataItem.Visible = true;
					groQG.Visible = true;
					this.Height = 488;
					break;
			}

			return( this.ShowDialog() );
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(ConfirmCopyForm));
			this.groEform = new System.Windows.Forms.GroupBox();
			this.chkEformTitle = new System.Windows.Forms.CheckBox();
			this.chkEformWidth = new System.Windows.Forms.CheckBox();
			this.chkEformHideIfInactive = new System.Windows.Forms.CheckBox();
			this.chkEformDisplayNumbers = new System.Windows.Forms.CheckBox();
			this.chkEformLocalLabel = new System.Windows.Forms.CheckBox();
			this.chkEformLabel = new System.Windows.Forms.CheckBox();
			this.chkEformBGColour = new System.Windows.Forms.CheckBox();
			this.btnNo = new System.Windows.Forms.Button();
			this.btnYes = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.groElement = new System.Windows.Forms.GroupBox();
			this.chkElementDisplayLength = new System.Windows.Forms.CheckBox();
			this.chkElementUse = new System.Windows.Forms.CheckBox();
			this.chkElementCaptionFontColour = new System.Windows.Forms.CheckBox();
			this.chkElementCaptionFontSize = new System.Windows.Forms.CheckBox();
			this.chkElementCaptionFontItalic = new System.Windows.Forms.CheckBox();
			this.chkElementCaptionFontBold = new System.Windows.Forms.CheckBox();
			this.chkElementCaptionFontName = new System.Windows.Forms.CheckBox();
			this.chkElementShowStatusFlag = new System.Windows.Forms.CheckBox();
			this.chkElementQGroupFieldOrder = new System.Windows.Forms.CheckBox();
			this.chkElementQGroupId = new System.Windows.Forms.CheckBox();
			this.chkElementOwnerQGroupId = new System.Windows.Forms.CheckBox();
			this.chkElementRoleCode = new System.Windows.Forms.CheckBox();
			this.chkElementRequireComment = new System.Windows.Forms.CheckBox();
			this.chkElementMandatory = new System.Windows.Forms.CheckBox();
			this.chkElementOptional = new System.Windows.Forms.CheckBox();
			this.chkElementLocalFlag = new System.Windows.Forms.CheckBox();
			this.chkElementHidden = new System.Windows.Forms.CheckBox();
			this.chkElementPrintOrder = new System.Windows.Forms.CheckBox();
			this.chkElementY = new System.Windows.Forms.CheckBox();
			this.chkElementX = new System.Windows.Forms.CheckBox();
			this.chkElementCaptionY = new System.Windows.Forms.CheckBox();
			this.chkElementCaptionX = new System.Windows.Forms.CheckBox();
			this.chkElementWidth = new System.Windows.Forms.CheckBox();
			this.chkElementHeight = new System.Windows.Forms.CheckBox();
			this.chkElementSkipCondition = new System.Windows.Forms.CheckBox();
			this.chkElementFieldOrder = new System.Windows.Forms.CheckBox();
			this.chkElementFontSize = new System.Windows.Forms.CheckBox();
			this.chkElementFontItalic = new System.Windows.Forms.CheckBox();
			this.chkElementFontBold = new System.Windows.Forms.CheckBox();
			this.chkElementFontName = new System.Windows.Forms.CheckBox();
			this.chkElementCaption = new System.Windows.Forms.CheckBox();
			this.chkElementFontColour = new System.Windows.Forms.CheckBox();
			this.chkElementControlType = new System.Windows.Forms.CheckBox();
			this.groDataItem = new System.Windows.Forms.GroupBox();
			this.chkDataItemValidation = new System.Windows.Forms.CheckBox();
			this.chkDataItemDescription = new System.Windows.Forms.CheckBox();
			this.chkDataItemCase = new System.Windows.Forms.CheckBox();
			this.chkDataItemHelpText = new System.Windows.Forms.CheckBox();
			this.chkDataItemDerivation = new System.Windows.Forms.CheckBox();
			this.chkDataItemLength = new System.Windows.Forms.CheckBox();
			this.chkDataItemUnitOfMeasurement = new System.Windows.Forms.CheckBox();
			this.chkDataItemFormat = new System.Windows.Forms.CheckBox();
			this.chkDataItemDataType = new System.Windows.Forms.CheckBox();
			this.groQG = new System.Windows.Forms.GroupBox();
			this.chkQGMaxRepeats = new System.Windows.Forms.CheckBox();
			this.chkQGMinRepeats = new System.Windows.Forms.CheckBox();
			this.chkQGInitialRows = new System.Windows.Forms.CheckBox();
			this.chkQGDisplayRows = new System.Windows.Forms.CheckBox();
			this.chkQGBorder = new System.Windows.Forms.CheckBox();
			this.chkQGDisplayType = new System.Windows.Forms.CheckBox();
			this.chkQGName = new System.Windows.Forms.CheckBox();
			this.groEform.SuspendLayout();
			this.groElement.SuspendLayout();
			this.groDataItem.SuspendLayout();
			this.groQG.SuspendLayout();
			this.SuspendLayout();
			// 
			// groEform
			// 
			this.groEform.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groEform.Controls.Add(this.chkEformTitle);
			this.groEform.Controls.Add(this.chkEformWidth);
			this.groEform.Controls.Add(this.chkEformHideIfInactive);
			this.groEform.Controls.Add(this.chkEformDisplayNumbers);
			this.groEform.Controls.Add(this.chkEformLocalLabel);
			this.groEform.Controls.Add(this.chkEformLabel);
			this.groEform.Controls.Add(this.chkEformBGColour);
			this.groEform.Location = new System.Drawing.Point(4, 48);
			this.groEform.Name = "groEform";
			this.groEform.Size = new System.Drawing.Size(576, 364);
			this.groEform.TabIndex = 0;
			this.groEform.TabStop = false;
			this.groEform.Text = "eForm Properties";
			// 
			// chkEformTitle
			// 
			this.chkEformTitle.Checked = true;
			this.chkEformTitle.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkEformTitle.Location = new System.Drawing.Point(428, 24);
			this.chkEformTitle.Name = "chkEformTitle";
			this.chkEformTitle.Size = new System.Drawing.Size(144, 16);
			this.chkEformTitle.TabIndex = 6;
			this.chkEformTitle.Text = "CRF Title";
			// 
			// chkEformWidth
			// 
			this.chkEformWidth.Checked = true;
			this.chkEformWidth.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkEformWidth.Location = new System.Drawing.Point(288, 44);
			this.chkEformWidth.Name = "chkEformWidth";
			this.chkEformWidth.Size = new System.Drawing.Size(144, 16);
			this.chkEformWidth.TabIndex = 5;
			this.chkEformWidth.Text = "eForm Width";
			// 
			// chkEformHideIfInactive
			// 
			this.chkEformHideIfInactive.Checked = true;
			this.chkEformHideIfInactive.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkEformHideIfInactive.Location = new System.Drawing.Point(148, 44);
			this.chkEformHideIfInactive.Name = "chkEformHideIfInactive";
			this.chkEformHideIfInactive.Size = new System.Drawing.Size(144, 16);
			this.chkEformHideIfInactive.TabIndex = 4;
			this.chkEformHideIfInactive.Text = "Hide If Inactive";
			// 
			// chkEformDisplayNumbers
			// 
			this.chkEformDisplayNumbers.Checked = true;
			this.chkEformDisplayNumbers.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkEformDisplayNumbers.Location = new System.Drawing.Point(8, 44);
			this.chkEformDisplayNumbers.Name = "chkEformDisplayNumbers";
			this.chkEformDisplayNumbers.Size = new System.Drawing.Size(144, 16);
			this.chkEformDisplayNumbers.TabIndex = 3;
			this.chkEformDisplayNumbers.Text = "Display Numbers";
			// 
			// chkEformLocalLabel
			// 
			this.chkEformLocalLabel.Checked = true;
			this.chkEformLocalLabel.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkEformLocalLabel.Location = new System.Drawing.Point(288, 24);
			this.chkEformLocalLabel.Name = "chkEformLocalLabel";
			this.chkEformLocalLabel.Size = new System.Drawing.Size(144, 16);
			this.chkEformLocalLabel.TabIndex = 2;
			this.chkEformLocalLabel.Text = "Local CRF Page Label";
			// 
			// chkEformLabel
			// 
			this.chkEformLabel.Checked = true;
			this.chkEformLabel.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkEformLabel.Location = new System.Drawing.Point(148, 24);
			this.chkEformLabel.Name = "chkEformLabel";
			this.chkEformLabel.Size = new System.Drawing.Size(144, 16);
			this.chkEformLabel.TabIndex = 1;
			this.chkEformLabel.Text = "CRF Page Label";
			// 
			// chkEformBGColour
			// 
			this.chkEformBGColour.Checked = true;
			this.chkEformBGColour.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkEformBGColour.Location = new System.Drawing.Point(8, 24);
			this.chkEformBGColour.Name = "chkEformBGColour";
			this.chkEformBGColour.Size = new System.Drawing.Size(144, 16);
			this.chkEformBGColour.TabIndex = 0;
			this.chkEformBGColour.Text = "Background Colour";
			// 
			// btnNo
			// 
			this.btnNo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnNo.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnNo.Location = new System.Drawing.Point(500, 424);
			this.btnNo.Name = "btnNo";
			this.btnNo.Size = new System.Drawing.Size(72, 24);
			this.btnNo.TabIndex = 1;
			this.btnNo.Text = "No";
			this.btnNo.Click += new System.EventHandler(this.btnNo_Click);
			// 
			// btnYes
			// 
			this.btnYes.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnYes.Location = new System.Drawing.Point(416, 424);
			this.btnYes.Name = "btnYes";
			this.btnYes.Size = new System.Drawing.Size(72, 24);
			this.btnYes.TabIndex = 2;
			this.btnYes.Text = "Yes";
			this.btnYes.Click += new System.EventHandler(this.btnYes_Click);
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(12, 12);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(448, 20);
			this.label1.TabIndex = 3;
			this.label1.Text = "Are you sure you want to copy the following properties for the selected matched i" +
				"tem(s)?";
			// 
			// groElement
			// 
			this.groElement.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groElement.Controls.Add(this.chkElementDisplayLength);
			this.groElement.Controls.Add(this.chkElementUse);
			this.groElement.Controls.Add(this.chkElementCaptionFontColour);
			this.groElement.Controls.Add(this.chkElementCaptionFontSize);
			this.groElement.Controls.Add(this.chkElementCaptionFontItalic);
			this.groElement.Controls.Add(this.chkElementCaptionFontBold);
			this.groElement.Controls.Add(this.chkElementCaptionFontName);
			this.groElement.Controls.Add(this.chkElementShowStatusFlag);
			this.groElement.Controls.Add(this.chkElementQGroupFieldOrder);
			this.groElement.Controls.Add(this.chkElementQGroupId);
			this.groElement.Controls.Add(this.chkElementOwnerQGroupId);
			this.groElement.Controls.Add(this.chkElementRoleCode);
			this.groElement.Controls.Add(this.chkElementRequireComment);
			this.groElement.Controls.Add(this.chkElementMandatory);
			this.groElement.Controls.Add(this.chkElementOptional);
			this.groElement.Controls.Add(this.chkElementLocalFlag);
			this.groElement.Controls.Add(this.chkElementHidden);
			this.groElement.Controls.Add(this.chkElementPrintOrder);
			this.groElement.Controls.Add(this.chkElementY);
			this.groElement.Controls.Add(this.chkElementX);
			this.groElement.Controls.Add(this.chkElementCaptionY);
			this.groElement.Controls.Add(this.chkElementCaptionX);
			this.groElement.Controls.Add(this.chkElementWidth);
			this.groElement.Controls.Add(this.chkElementHeight);
			this.groElement.Controls.Add(this.chkElementSkipCondition);
			this.groElement.Controls.Add(this.chkElementFieldOrder);
			this.groElement.Controls.Add(this.chkElementFontSize);
			this.groElement.Controls.Add(this.chkElementFontItalic);
			this.groElement.Controls.Add(this.chkElementFontBold);
			this.groElement.Controls.Add(this.chkElementFontName);
			this.groElement.Controls.Add(this.chkElementCaption);
			this.groElement.Controls.Add(this.chkElementFontColour);
			this.groElement.Controls.Add(this.chkElementControlType);
			this.groElement.Location = new System.Drawing.Point(4, 48);
			this.groElement.Name = "groElement";
			this.groElement.Size = new System.Drawing.Size(576, 196);
			this.groElement.TabIndex = 4;
			this.groElement.TabStop = false;
			this.groElement.Text = "eForm Element Properties";
			// 
			// chkElementDisplayLength
			// 
			this.chkElementDisplayLength.Checked = true;
			this.chkElementDisplayLength.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkElementDisplayLength.Location = new System.Drawing.Point(428, 64);
			this.chkElementDisplayLength.Name = "chkElementDisplayLength";
			this.chkElementDisplayLength.Size = new System.Drawing.Size(144, 16);
			this.chkElementDisplayLength.TabIndex = 32;
			this.chkElementDisplayLength.Text = "Display Length";
			// 
			// chkElementUse
			// 
			this.chkElementUse.Checked = true;
			this.chkElementUse.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkElementUse.Location = new System.Drawing.Point(428, 118);
			this.chkElementUse.Name = "chkElementUse";
			this.chkElementUse.Size = new System.Drawing.Size(144, 16);
			this.chkElementUse.TabIndex = 31;
			this.chkElementUse.Text = "Element Use";
			// 
			// chkElementCaptionFontColour
			// 
			this.chkElementCaptionFontColour.Checked = true;
			this.chkElementCaptionFontColour.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkElementCaptionFontColour.Location = new System.Drawing.Point(428, 154);
			this.chkElementCaptionFontColour.Name = "chkElementCaptionFontColour";
			this.chkElementCaptionFontColour.Size = new System.Drawing.Size(144, 16);
			this.chkElementCaptionFontColour.TabIndex = 30;
			this.chkElementCaptionFontColour.Text = "Caption Font Colour";
			// 
			// chkElementCaptionFontSize
			// 
			this.chkElementCaptionFontSize.Checked = true;
			this.chkElementCaptionFontSize.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkElementCaptionFontSize.Location = new System.Drawing.Point(428, 46);
			this.chkElementCaptionFontSize.Name = "chkElementCaptionFontSize";
			this.chkElementCaptionFontSize.Size = new System.Drawing.Size(144, 16);
			this.chkElementCaptionFontSize.TabIndex = 29;
			this.chkElementCaptionFontSize.Text = "Caption Font Size";
			// 
			// chkElementCaptionFontItalic
			// 
			this.chkElementCaptionFontItalic.Checked = true;
			this.chkElementCaptionFontItalic.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkElementCaptionFontItalic.Location = new System.Drawing.Point(428, 100);
			this.chkElementCaptionFontItalic.Name = "chkElementCaptionFontItalic";
			this.chkElementCaptionFontItalic.Size = new System.Drawing.Size(144, 16);
			this.chkElementCaptionFontItalic.TabIndex = 28;
			this.chkElementCaptionFontItalic.Text = "Caption Font Italic";
			// 
			// chkElementCaptionFontBold
			// 
			this.chkElementCaptionFontBold.Checked = true;
			this.chkElementCaptionFontBold.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkElementCaptionFontBold.Location = new System.Drawing.Point(428, 136);
			this.chkElementCaptionFontBold.Name = "chkElementCaptionFontBold";
			this.chkElementCaptionFontBold.Size = new System.Drawing.Size(144, 16);
			this.chkElementCaptionFontBold.TabIndex = 27;
			this.chkElementCaptionFontBold.Text = "Caption Font Bold";
			// 
			// chkElementCaptionFontName
			// 
			this.chkElementCaptionFontName.Checked = true;
			this.chkElementCaptionFontName.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkElementCaptionFontName.Location = new System.Drawing.Point(428, 28);
			this.chkElementCaptionFontName.Name = "chkElementCaptionFontName";
			this.chkElementCaptionFontName.Size = new System.Drawing.Size(144, 16);
			this.chkElementCaptionFontName.TabIndex = 26;
			this.chkElementCaptionFontName.Text = "Caption Font Name";
			// 
			// chkElementShowStatusFlag
			// 
			this.chkElementShowStatusFlag.Checked = true;
			this.chkElementShowStatusFlag.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkElementShowStatusFlag.Location = new System.Drawing.Point(428, 82);
			this.chkElementShowStatusFlag.Name = "chkElementShowStatusFlag";
			this.chkElementShowStatusFlag.Size = new System.Drawing.Size(144, 16);
			this.chkElementShowStatusFlag.TabIndex = 25;
			this.chkElementShowStatusFlag.Text = "Show Status Flag";
			// 
			// chkElementQGroupFieldOrder
			// 
			this.chkElementQGroupFieldOrder.Enabled = false;
			this.chkElementQGroupFieldOrder.Location = new System.Drawing.Point(8, 172);
			this.chkElementQGroupFieldOrder.Name = "chkElementQGroupFieldOrder";
			this.chkElementQGroupFieldOrder.Size = new System.Drawing.Size(144, 16);
			this.chkElementQGroupFieldOrder.TabIndex = 24;
			this.chkElementQGroupFieldOrder.Text = "Q Group Field Order";
			// 
			// chkElementQGroupId
			// 
			this.chkElementQGroupId.Enabled = false;
			this.chkElementQGroupId.Location = new System.Drawing.Point(288, 154);
			this.chkElementQGroupId.Name = "chkElementQGroupId";
			this.chkElementQGroupId.Size = new System.Drawing.Size(144, 16);
			this.chkElementQGroupId.TabIndex = 23;
			this.chkElementQGroupId.Text = "Q Group Id";
			// 
			// chkElementOwnerQGroupId
			// 
			this.chkElementOwnerQGroupId.Enabled = false;
			this.chkElementOwnerQGroupId.Location = new System.Drawing.Point(148, 154);
			this.chkElementOwnerQGroupId.Name = "chkElementOwnerQGroupId";
			this.chkElementOwnerQGroupId.Size = new System.Drawing.Size(144, 16);
			this.chkElementOwnerQGroupId.TabIndex = 22;
			this.chkElementOwnerQGroupId.Text = "Owner Q Group Id";
			// 
			// chkElementRoleCode
			// 
			this.chkElementRoleCode.Checked = true;
			this.chkElementRoleCode.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkElementRoleCode.Location = new System.Drawing.Point(8, 154);
			this.chkElementRoleCode.Name = "chkElementRoleCode";
			this.chkElementRoleCode.Size = new System.Drawing.Size(144, 16);
			this.chkElementRoleCode.TabIndex = 21;
			this.chkElementRoleCode.Text = "Role Code";
			// 
			// chkElementRequireComment
			// 
			this.chkElementRequireComment.Checked = true;
			this.chkElementRequireComment.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkElementRequireComment.Location = new System.Drawing.Point(288, 136);
			this.chkElementRequireComment.Name = "chkElementRequireComment";
			this.chkElementRequireComment.Size = new System.Drawing.Size(144, 16);
			this.chkElementRequireComment.TabIndex = 20;
			this.chkElementRequireComment.Text = "Require Comment";
			// 
			// chkElementMandatory
			// 
			this.chkElementMandatory.Checked = true;
			this.chkElementMandatory.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkElementMandatory.Location = new System.Drawing.Point(148, 136);
			this.chkElementMandatory.Name = "chkElementMandatory";
			this.chkElementMandatory.Size = new System.Drawing.Size(144, 16);
			this.chkElementMandatory.TabIndex = 19;
			this.chkElementMandatory.Text = "Mandatory";
			// 
			// chkElementOptional
			// 
			this.chkElementOptional.Checked = true;
			this.chkElementOptional.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkElementOptional.Location = new System.Drawing.Point(8, 136);
			this.chkElementOptional.Name = "chkElementOptional";
			this.chkElementOptional.Size = new System.Drawing.Size(144, 16);
			this.chkElementOptional.TabIndex = 18;
			this.chkElementOptional.Text = "Optional";
			// 
			// chkElementLocalFlag
			// 
			this.chkElementLocalFlag.Checked = true;
			this.chkElementLocalFlag.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkElementLocalFlag.Location = new System.Drawing.Point(288, 118);
			this.chkElementLocalFlag.Name = "chkElementLocalFlag";
			this.chkElementLocalFlag.Size = new System.Drawing.Size(144, 16);
			this.chkElementLocalFlag.TabIndex = 17;
			this.chkElementLocalFlag.Text = "Local Flag";
			// 
			// chkElementHidden
			// 
			this.chkElementHidden.Checked = true;
			this.chkElementHidden.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkElementHidden.Location = new System.Drawing.Point(148, 118);
			this.chkElementHidden.Name = "chkElementHidden";
			this.chkElementHidden.Size = new System.Drawing.Size(144, 16);
			this.chkElementHidden.TabIndex = 16;
			this.chkElementHidden.Text = "Hidden";
			// 
			// chkElementPrintOrder
			// 
			this.chkElementPrintOrder.Checked = true;
			this.chkElementPrintOrder.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkElementPrintOrder.Location = new System.Drawing.Point(8, 118);
			this.chkElementPrintOrder.Name = "chkElementPrintOrder";
			this.chkElementPrintOrder.Size = new System.Drawing.Size(144, 16);
			this.chkElementPrintOrder.TabIndex = 15;
			this.chkElementPrintOrder.Text = "Print Order";
			// 
			// chkElementY
			// 
			this.chkElementY.Checked = true;
			this.chkElementY.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkElementY.Location = new System.Drawing.Point(288, 100);
			this.chkElementY.Name = "chkElementY";
			this.chkElementY.Size = new System.Drawing.Size(144, 16);
			this.chkElementY.TabIndex = 14;
			this.chkElementY.Text = "Y";
			// 
			// chkElementX
			// 
			this.chkElementX.Checked = true;
			this.chkElementX.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkElementX.Location = new System.Drawing.Point(148, 100);
			this.chkElementX.Name = "chkElementX";
			this.chkElementX.Size = new System.Drawing.Size(144, 16);
			this.chkElementX.TabIndex = 13;
			this.chkElementX.Text = "X";
			// 
			// chkElementCaptionY
			// 
			this.chkElementCaptionY.Checked = true;
			this.chkElementCaptionY.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkElementCaptionY.Location = new System.Drawing.Point(8, 100);
			this.chkElementCaptionY.Name = "chkElementCaptionY";
			this.chkElementCaptionY.Size = new System.Drawing.Size(144, 16);
			this.chkElementCaptionY.TabIndex = 12;
			this.chkElementCaptionY.Text = "Caption Y";
			// 
			// chkElementCaptionX
			// 
			this.chkElementCaptionX.Checked = true;
			this.chkElementCaptionX.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkElementCaptionX.Location = new System.Drawing.Point(288, 82);
			this.chkElementCaptionX.Name = "chkElementCaptionX";
			this.chkElementCaptionX.Size = new System.Drawing.Size(144, 16);
			this.chkElementCaptionX.TabIndex = 11;
			this.chkElementCaptionX.Text = "Caption X";
			// 
			// chkElementWidth
			// 
			this.chkElementWidth.Checked = true;
			this.chkElementWidth.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkElementWidth.Location = new System.Drawing.Point(148, 82);
			this.chkElementWidth.Name = "chkElementWidth";
			this.chkElementWidth.Size = new System.Drawing.Size(144, 16);
			this.chkElementWidth.TabIndex = 10;
			this.chkElementWidth.Text = "Width";
			// 
			// chkElementHeight
			// 
			this.chkElementHeight.Checked = true;
			this.chkElementHeight.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkElementHeight.Location = new System.Drawing.Point(8, 82);
			this.chkElementHeight.Name = "chkElementHeight";
			this.chkElementHeight.Size = new System.Drawing.Size(144, 16);
			this.chkElementHeight.TabIndex = 9;
			this.chkElementHeight.Text = "Height";
			// 
			// chkElementSkipCondition
			// 
			this.chkElementSkipCondition.Checked = true;
			this.chkElementSkipCondition.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkElementSkipCondition.Location = new System.Drawing.Point(288, 64);
			this.chkElementSkipCondition.Name = "chkElementSkipCondition";
			this.chkElementSkipCondition.Size = new System.Drawing.Size(144, 16);
			this.chkElementSkipCondition.TabIndex = 8;
			this.chkElementSkipCondition.Text = "Skip Condition";
			// 
			// chkElementFieldOrder
			// 
			this.chkElementFieldOrder.Enabled = false;
			this.chkElementFieldOrder.Location = new System.Drawing.Point(148, 64);
			this.chkElementFieldOrder.Name = "chkElementFieldOrder";
			this.chkElementFieldOrder.Size = new System.Drawing.Size(144, 16);
			this.chkElementFieldOrder.TabIndex = 7;
			this.chkElementFieldOrder.Text = "Field Order";
			// 
			// chkElementFontSize
			// 
			this.chkElementFontSize.Checked = true;
			this.chkElementFontSize.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkElementFontSize.Location = new System.Drawing.Point(8, 64);
			this.chkElementFontSize.Name = "chkElementFontSize";
			this.chkElementFontSize.Size = new System.Drawing.Size(144, 16);
			this.chkElementFontSize.TabIndex = 6;
			this.chkElementFontSize.Text = "Font Size";
			// 
			// chkElementFontItalic
			// 
			this.chkElementFontItalic.Checked = true;
			this.chkElementFontItalic.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkElementFontItalic.Location = new System.Drawing.Point(288, 46);
			this.chkElementFontItalic.Name = "chkElementFontItalic";
			this.chkElementFontItalic.Size = new System.Drawing.Size(144, 16);
			this.chkElementFontItalic.TabIndex = 5;
			this.chkElementFontItalic.Text = "Font Italic";
			// 
			// chkElementFontBold
			// 
			this.chkElementFontBold.Checked = true;
			this.chkElementFontBold.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkElementFontBold.Location = new System.Drawing.Point(148, 46);
			this.chkElementFontBold.Name = "chkElementFontBold";
			this.chkElementFontBold.Size = new System.Drawing.Size(144, 16);
			this.chkElementFontBold.TabIndex = 4;
			this.chkElementFontBold.Text = "Font Bold";
			// 
			// chkElementFontName
			// 
			this.chkElementFontName.Checked = true;
			this.chkElementFontName.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkElementFontName.Location = new System.Drawing.Point(8, 46);
			this.chkElementFontName.Name = "chkElementFontName";
			this.chkElementFontName.Size = new System.Drawing.Size(144, 16);
			this.chkElementFontName.TabIndex = 3;
			this.chkElementFontName.Text = "Font Name";
			// 
			// chkElementCaption
			// 
			this.chkElementCaption.Checked = true;
			this.chkElementCaption.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkElementCaption.Location = new System.Drawing.Point(288, 28);
			this.chkElementCaption.Name = "chkElementCaption";
			this.chkElementCaption.Size = new System.Drawing.Size(144, 16);
			this.chkElementCaption.TabIndex = 2;
			this.chkElementCaption.Text = "Caption";
			// 
			// chkElementFontColour
			// 
			this.chkElementFontColour.Checked = true;
			this.chkElementFontColour.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkElementFontColour.Location = new System.Drawing.Point(148, 28);
			this.chkElementFontColour.Name = "chkElementFontColour";
			this.chkElementFontColour.Size = new System.Drawing.Size(144, 16);
			this.chkElementFontColour.TabIndex = 1;
			this.chkElementFontColour.Text = "Font Colour";
			// 
			// chkElementControlType
			// 
			this.chkElementControlType.Checked = true;
			this.chkElementControlType.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkElementControlType.Location = new System.Drawing.Point(8, 28);
			this.chkElementControlType.Name = "chkElementControlType";
			this.chkElementControlType.Size = new System.Drawing.Size(144, 16);
			this.chkElementControlType.TabIndex = 0;
			this.chkElementControlType.Text = "Control Type";
			// 
			// groDataItem
			// 
			this.groDataItem.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groDataItem.Controls.Add(this.chkDataItemValidation);
			this.groDataItem.Controls.Add(this.chkDataItemDescription);
			this.groDataItem.Controls.Add(this.chkDataItemCase);
			this.groDataItem.Controls.Add(this.chkDataItemHelpText);
			this.groDataItem.Controls.Add(this.chkDataItemDerivation);
			this.groDataItem.Controls.Add(this.chkDataItemLength);
			this.groDataItem.Controls.Add(this.chkDataItemUnitOfMeasurement);
			this.groDataItem.Controls.Add(this.chkDataItemFormat);
			this.groDataItem.Controls.Add(this.chkDataItemDataType);
			this.groDataItem.Location = new System.Drawing.Point(5, 252);
			this.groDataItem.Name = "groDataItem";
			this.groDataItem.Size = new System.Drawing.Size(576, 88);
			this.groDataItem.TabIndex = 10;
			this.groDataItem.TabStop = false;
			this.groDataItem.Text = "DataItem Properties";
			// 
			// chkDataItemValidation
			// 
			this.chkDataItemValidation.Checked = true;
			this.chkDataItemValidation.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkDataItemValidation.Location = new System.Drawing.Point(428, 24);
			this.chkDataItemValidation.Name = "chkDataItemValidation";
			this.chkDataItemValidation.Size = new System.Drawing.Size(144, 16);
			this.chkDataItemValidation.TabIndex = 8;
			this.chkDataItemValidation.Text = "DataItem Validation";
			// 
			// chkDataItemDescription
			// 
			this.chkDataItemDescription.Checked = true;
			this.chkDataItemDescription.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkDataItemDescription.Location = new System.Drawing.Point(428, 44);
			this.chkDataItemDescription.Name = "chkDataItemDescription";
			this.chkDataItemDescription.Size = new System.Drawing.Size(144, 16);
			this.chkDataItemDescription.TabIndex = 7;
			this.chkDataItemDescription.Text = "Description";
			// 
			// chkDataItemCase
			// 
			this.chkDataItemCase.Checked = true;
			this.chkDataItemCase.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkDataItemCase.Location = new System.Drawing.Point(8, 64);
			this.chkDataItemCase.Name = "chkDataItemCase";
			this.chkDataItemCase.Size = new System.Drawing.Size(144, 16);
			this.chkDataItemCase.TabIndex = 6;
			this.chkDataItemCase.Text = "DataItem Case";
			// 
			// chkDataItemHelpText
			// 
			this.chkDataItemHelpText.Checked = true;
			this.chkDataItemHelpText.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkDataItemHelpText.Location = new System.Drawing.Point(288, 44);
			this.chkDataItemHelpText.Name = "chkDataItemHelpText";
			this.chkDataItemHelpText.Size = new System.Drawing.Size(144, 16);
			this.chkDataItemHelpText.TabIndex = 5;
			this.chkDataItemHelpText.Text = "DataItem Help Text";
			// 
			// chkDataItemDerivation
			// 
			this.chkDataItemDerivation.Checked = true;
			this.chkDataItemDerivation.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkDataItemDerivation.Location = new System.Drawing.Point(148, 44);
			this.chkDataItemDerivation.Name = "chkDataItemDerivation";
			this.chkDataItemDerivation.Size = new System.Drawing.Size(144, 16);
			this.chkDataItemDerivation.TabIndex = 4;
			this.chkDataItemDerivation.Text = "Derivation";
			// 
			// chkDataItemLength
			// 
			this.chkDataItemLength.Checked = true;
			this.chkDataItemLength.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkDataItemLength.Location = new System.Drawing.Point(8, 44);
			this.chkDataItemLength.Name = "chkDataItemLength";
			this.chkDataItemLength.Size = new System.Drawing.Size(144, 16);
			this.chkDataItemLength.TabIndex = 3;
			this.chkDataItemLength.Text = "DataItem Length";
			// 
			// chkDataItemUnitOfMeasurement
			// 
			this.chkDataItemUnitOfMeasurement.Checked = true;
			this.chkDataItemUnitOfMeasurement.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkDataItemUnitOfMeasurement.Location = new System.Drawing.Point(288, 24);
			this.chkDataItemUnitOfMeasurement.Name = "chkDataItemUnitOfMeasurement";
			this.chkDataItemUnitOfMeasurement.Size = new System.Drawing.Size(144, 16);
			this.chkDataItemUnitOfMeasurement.TabIndex = 2;
			this.chkDataItemUnitOfMeasurement.Text = "Unit Of Measurement";
			// 
			// chkDataItemFormat
			// 
			this.chkDataItemFormat.Checked = true;
			this.chkDataItemFormat.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkDataItemFormat.Location = new System.Drawing.Point(148, 24);
			this.chkDataItemFormat.Name = "chkDataItemFormat";
			this.chkDataItemFormat.Size = new System.Drawing.Size(144, 16);
			this.chkDataItemFormat.TabIndex = 1;
			this.chkDataItemFormat.Text = "DataItem Format";
			// 
			// chkDataItemDataType
			// 
			this.chkDataItemDataType.Enabled = false;
			this.chkDataItemDataType.Location = new System.Drawing.Point(8, 24);
			this.chkDataItemDataType.Name = "chkDataItemDataType";
			this.chkDataItemDataType.Size = new System.Drawing.Size(144, 16);
			this.chkDataItemDataType.TabIndex = 0;
			this.chkDataItemDataType.Text = "DataType";
			// 
			// groQG
			// 
			this.groQG.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groQG.Controls.Add(this.chkQGMaxRepeats);
			this.groQG.Controls.Add(this.chkQGMinRepeats);
			this.groQG.Controls.Add(this.chkQGInitialRows);
			this.groQG.Controls.Add(this.chkQGDisplayRows);
			this.groQG.Controls.Add(this.chkQGBorder);
			this.groQG.Controls.Add(this.chkQGDisplayType);
			this.groQG.Controls.Add(this.chkQGName);
			this.groQG.Location = new System.Drawing.Point(4, 344);
			this.groQG.Name = "groQG";
			this.groQG.Size = new System.Drawing.Size(576, 68);
			this.groQG.TabIndex = 11;
			this.groQG.TabStop = false;
			this.groQG.Text = "Question Group Properties";
			// 
			// chkQGMaxRepeats
			// 
			this.chkQGMaxRepeats.Checked = true;
			this.chkQGMaxRepeats.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkQGMaxRepeats.Location = new System.Drawing.Point(288, 44);
			this.chkQGMaxRepeats.Name = "chkQGMaxRepeats";
			this.chkQGMaxRepeats.Size = new System.Drawing.Size(144, 16);
			this.chkQGMaxRepeats.TabIndex = 6;
			this.chkQGMaxRepeats.Text = "Max. Repeats";
			// 
			// chkQGMinRepeats
			// 
			this.chkQGMinRepeats.Checked = true;
			this.chkQGMinRepeats.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkQGMinRepeats.Location = new System.Drawing.Point(148, 44);
			this.chkQGMinRepeats.Name = "chkQGMinRepeats";
			this.chkQGMinRepeats.Size = new System.Drawing.Size(144, 16);
			this.chkQGMinRepeats.TabIndex = 5;
			this.chkQGMinRepeats.Text = "Min. Repeats";
			// 
			// chkQGInitialRows
			// 
			this.chkQGInitialRows.Checked = true;
			this.chkQGInitialRows.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkQGInitialRows.Location = new System.Drawing.Point(8, 44);
			this.chkQGInitialRows.Name = "chkQGInitialRows";
			this.chkQGInitialRows.Size = new System.Drawing.Size(144, 16);
			this.chkQGInitialRows.TabIndex = 4;
			this.chkQGInitialRows.Text = "Initial Rows";
			// 
			// chkQGDisplayRows
			// 
			this.chkQGDisplayRows.Checked = true;
			this.chkQGDisplayRows.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkQGDisplayRows.Location = new System.Drawing.Point(428, 24);
			this.chkQGDisplayRows.Name = "chkQGDisplayRows";
			this.chkQGDisplayRows.Size = new System.Drawing.Size(144, 16);
			this.chkQGDisplayRows.TabIndex = 3;
			this.chkQGDisplayRows.Text = "Display Rows";
			// 
			// chkQGBorder
			// 
			this.chkQGBorder.Checked = true;
			this.chkQGBorder.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkQGBorder.Location = new System.Drawing.Point(288, 24);
			this.chkQGBorder.Name = "chkQGBorder";
			this.chkQGBorder.Size = new System.Drawing.Size(144, 16);
			this.chkQGBorder.TabIndex = 2;
			this.chkQGBorder.Text = "Border";
			// 
			// chkQGDisplayType
			// 
			this.chkQGDisplayType.Checked = true;
			this.chkQGDisplayType.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkQGDisplayType.Location = new System.Drawing.Point(148, 24);
			this.chkQGDisplayType.Name = "chkQGDisplayType";
			this.chkQGDisplayType.Size = new System.Drawing.Size(144, 16);
			this.chkQGDisplayType.TabIndex = 1;
			this.chkQGDisplayType.Text = "Display Type";
			// 
			// chkQGName
			// 
			this.chkQGName.Checked = true;
			this.chkQGName.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkQGName.Location = new System.Drawing.Point(8, 24);
			this.chkQGName.Name = "chkQGName";
			this.chkQGName.Size = new System.Drawing.Size(144, 16);
			this.chkQGName.TabIndex = 0;
			this.chkQGName.Text = "QGroup Name";
			// 
			// ConfirmCopyForm
			// 
			this.AcceptButton = this.btnYes;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CancelButton = this.btnNo;
			this.ClientSize = new System.Drawing.Size(582, 456);
			this.Controls.Add(this.groDataItem);
			this.Controls.Add(this.groQG);
			this.Controls.Add(this.groElement);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btnYes);
			this.Controls.Add(this.btnNo);
			this.Controls.Add(this.groEform);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "ConfirmCopyForm";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Confirm Copy Properties";
			this.groEform.ResumeLayout(false);
			this.groElement.ResumeLayout(false);
			this.groDataItem.ResumeLayout(false);
			this.groQG.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void btnYes_Click(object sender, System.EventArgs e)
		{
			switch( _eType )
			{
				case StudyCopyGlobal.ElementType.eForm:
					if( chkEformBGColour.Checked )_elementCopyProperties.Add( "BACKGROUNDCOLOUR" );
					if( chkEformLabel.Checked ) _elementCopyProperties.Add( "CRFPAGELABEL" );
					if( chkEformLocalLabel.Checked ) _elementCopyProperties.Add( "LOCALCRFPAGELABEL" );
					if( chkEformDisplayNumbers.Checked ) _elementCopyProperties.Add ( "DISPLAYNUMBERS" );
					if( chkEformHideIfInactive.Checked ) _elementCopyProperties.Add( "HIDEIFINACTIVE" );
					if( chkEformWidth.Checked ) _elementCopyProperties.Add ( "EFORMWIDTH" );
					if( chkEformTitle.Checked ) _elementCopyProperties.Add( "CRFTITLE" );
					break;

				case StudyCopyGlobal.ElementType.eFormElement:
					//element
					if( chkElementControlType.Checked )_elementCopyProperties.Add( "CONTROLTYPE" );
					if( chkElementFontColour.Checked )_elementCopyProperties.Add( "FONTCOLOUR" );
					if( chkElementCaption.Checked )_elementCopyProperties.Add( "CAPTION" );
					if( chkElementFontName.Checked )_elementCopyProperties.Add( "FONTNAME" );
					if( chkElementFontBold.Checked )_elementCopyProperties.Add( "FONTBOLD" );
					if( chkElementFontItalic.Checked )_elementCopyProperties.Add( "FONTITALIC" );
					if( chkElementFontSize.Checked )_elementCopyProperties.Add( "FONTSIZE" );
					if( chkElementFieldOrder.Checked )_elementCopyProperties.Add( "FIELDORDER" );
					if( chkElementSkipCondition.Checked )_elementCopyProperties.Add( "SKIPCONDITION" );
					if( chkElementHeight.Checked )_elementCopyProperties.Add( "HEIGHT" );
					if( chkElementWidth.Checked )_elementCopyProperties.Add( "WIDTH" );
					if( chkElementCaptionX.Checked )_elementCopyProperties.Add( "CAPTIONX" );
					if( chkElementCaptionY.Checked )_elementCopyProperties.Add( "CAPTIONY" );
					if( chkElementX.Checked )_elementCopyProperties.Add( "X" );
					if( chkElementY.Checked )_elementCopyProperties.Add( "Y" );
					if( chkElementPrintOrder.Checked )_elementCopyProperties.Add( "PRINTORDER" );
					if( chkElementHidden.Checked )_elementCopyProperties.Add( "HIDDEN" );
					if( chkElementLocalFlag.Checked )_elementCopyProperties.Add( "LOCALFLAG" );
					if( chkElementOptional.Checked )_elementCopyProperties.Add( "OPTIONAL" );
					if( chkElementMandatory.Checked )_elementCopyProperties.Add( "MANDATORY" );
					if( chkElementRequireComment.Checked )_elementCopyProperties.Add( "REQUIRECOMMENT" );
					if( chkElementRoleCode.Checked )_elementCopyProperties.Add( "ROLECODE" );
					if( chkElementOwnerQGroupId.Checked )_elementCopyProperties.Add( "OWNERQGROUPID" );
					if( chkElementQGroupId.Checked )_elementCopyProperties.Add( "QGROUPID" );
					if( chkElementQGroupFieldOrder.Checked )_elementCopyProperties.Add( "QGROUPFIELDORDER" );
					if( chkElementShowStatusFlag.Checked )_elementCopyProperties.Add( "SHOWSTATUSFLAG" );
					if( chkElementCaptionFontName.Checked )_elementCopyProperties.Add( "CAPTIONFONTNAME" );
					if( chkElementCaptionFontBold.Checked )_elementCopyProperties.Add( "CAPTIONFONTBOLD" );
					if( chkElementCaptionFontItalic.Checked )_elementCopyProperties.Add( "CAPTIONFONTITALIC" );
					if( chkElementCaptionFontSize.Checked )_elementCopyProperties.Add( "CAPTIONFONTSIZE" );
					if( chkElementCaptionFontColour.Checked )_elementCopyProperties.Add( "CAPTIONFONTCOLOUR" );
					if( chkElementUse.Checked )_elementCopyProperties.Add( "ELEMENTUSE" );
					if( chkElementDisplayLength.Checked )_elementCopyProperties.Add( "DISPLAYLENGTH" );

					//dataitem
					if( chkDataItemDataType.Checked )_dataItemCopyProperties.Add( "DATATYPE" );
					if( chkDataItemFormat.Checked )_dataItemCopyProperties.Add( "DATAITEMFORMAT" );
					if( chkDataItemUnitOfMeasurement.Checked )_dataItemCopyProperties.Add( "UNITOFMEASUREMENT" );
					if( chkDataItemLength.Checked )_dataItemCopyProperties.Add( "DATAITEMLENGTH" );
					if( chkDataItemDerivation.Checked )_dataItemCopyProperties.Add( "DERIVATION" );
					if( chkDataItemHelpText.Checked )_dataItemCopyProperties.Add( "DATAITEMHELPTEXT" );
					if( chkDataItemCase.Checked )_dataItemCopyProperties.Add( "DATAITEMCASE" );
					if( chkDataItemDescription.Checked )_dataItemCopyProperties.Add( "DESCRIPTION" );

					//qg
					if( chkQGName.Checked )_qgCopyProperties.Add( "QGROUPNAME" );
					if( chkQGDisplayType.Checked )_qgCopyProperties.Add( "DISPLAYTYPE" );

					//eform qg
					if( chkQGBorder.Checked )_eformQGCopyProperties.Add( "BORDER" );
					if( chkQGDisplayRows.Checked )_eformQGCopyProperties.Add( "DISPLAYROWS" );
					if( chkQGInitialRows.Checked )_eformQGCopyProperties.Add( "INITIALROWS" );
					if( chkQGMinRepeats.Checked )_eformQGCopyProperties.Add( "MINREPEATS" );
					if( chkQGMaxRepeats.Checked )_eformQGCopyProperties.Add( "MAXREPEATS" );
				break;
			}

			if(
				((_eType == StudyCopyGlobal.ElementType.eForm) 
				&& (_elementCopyProperties.Count == 0))
				||
				((_eType != StudyCopyGlobal.ElementType.eForm) 
				&& ((_elementCopyProperties.Count == 0) 
				&& (_dataItemCopyProperties.Count == 0)
				&& (chkDataItemValidation.Checked == false)
				&& (_qgCopyProperties.Count == 0)
				&& (_eformQGCopyProperties.Count == 0))
				))
			{
				MessageBox.Show( "You must select at least one property to copy", "No Properties Selected", 
				MessageBoxButtons.OK, MessageBoxIcon.Information );
				_elementCopyProperties.Clear();
				_dataItemCopyProperties.Clear();
				_qgCopyProperties.Clear();
				_eformQGCopyProperties.Clear();
				return;
			}

			this.DialogResult = DialogResult.Yes;
			this.Close();
		}

		private void btnNo_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		public ArrayList ElementCopyProperties
		{
			get{ return( _elementCopyProperties ); }
		}

		public ArrayList DataItemCopyProperties
		{
			get{ return( _dataItemCopyProperties ); }
		}

		public ArrayList QGCopyProperties
		{
			get{ return( _qgCopyProperties ); }
		}

		public ArrayList EformQGCopyProperties
		{
			get{ return( _eformQGCopyProperties ); }
		}

		public bool DataItemValidation
		{
			get{ return( chkDataItemValidation.Checked ); }
		}

		private ArrayList _elementCopyProperties = new ArrayList();
		private ArrayList _dataItemCopyProperties = new ArrayList();
		private ArrayList _qgCopyProperties = new ArrayList();
		private ArrayList _eformQGCopyProperties = new ArrayList();
		private StudyCopyGlobal.ElementType _eType;
	}
}
