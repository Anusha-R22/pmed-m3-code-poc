using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using InferMed.MACROBuffer;
using System.Xml;
using System.IO;

namespace MACROBufferAPITestHarness
{
    public partial class DataEntryForm : Form
    {

        private string _dataXml;
        private string _resultXml;

        public DataEntryForm()
        {
            InitializeComponent();
            _dataXml = "";
            _resultXml = "";
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (chkVisitDate.Checked)
            {
                grpVisitDate.Enabled = true;
            }
            else
            {
                grpVisitDate.Enabled = false;
            }
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void chkCoverSheet_CheckedChanged(object sender, EventArgs e)
        {
            if (chkCoverSheet.Checked)
            {
                grpCoverSheet.Enabled = true;
            }
            else
            {
                grpCoverSheet.Enabled = false;
            }
        }

        private void chkInclusion_CheckedChanged(object sender, EventArgs e)
        {
            if (chkInclusion.Checked)
            {
                grpInclusion.Enabled = true;
            }
            else
            {
                grpInclusion.Enabled = false;
            }
        }

        private void chkDemographics_CheckedChanged(object sender, EventArgs e)
        {
            if (chkDemographics.Checked)
            {
                grpDemographics.Enabled = true;
            }
            else
            {
                grpDemographics.Enabled = false;
            }
        }

        private void chkAdverseEvent_CheckedChanged(object sender, EventArgs e)
        {
            if (chkAdverseEvent.Checked)
            {
                grpAdverseEvents.Enabled = true;
            }
            else
            {
                grpAdverseEvents.Enabled = false;
            }
        }

        private void DataEntryForm_Load(object sender, EventArgs e)
        {
            chkCoverSheet.Checked = false;
            chkDemographics.Checked = false;
            chkVisitDate.Checked = false;
            chkInclusion.Checked = false;
            chkAdverseEvent.Checked = false;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (ValidatePage())
            {
                // create CDISC XML
                string sCDISC = CreateCDISCXml();
                // store in last data
                _dataXml = sCDISC;
            
                // post to API
                MACROBufferAPI macroBufferApi = new MACROBufferAPI();
                string result = macroBufferApi.WriteBufferMessage(_dataXml, false);
                
                //string result = "";

                // handle results
                // write to textbox
                ParseResultsXML(result);

                // store in last results
                _resultXml = result;

            }
        }

        private string CreateCDISCXml()
        {
            StringBuilder CdiscXml = new StringBuilder();
            // header
            CdiscXml.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            CdiscXml.Append("<!DOCTYPE ODM SYSTEM \"");
            // local odm web version - "http://www.cdisc.org/dtd/ODM1-2-0.dtd"
            CdiscXml.Append(Application.StartupPath + "/ODM1-2-0.dtd");
            CdiscXml.Append("\">");
            CdiscXml.Append("<ODM FileType=\"Transactional\" FileOID=\"Test\" CreationDateTime=\"");
            CdiscXml.Append(DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss"));
            CdiscXml.Append("\">");
            CdiscXml.Append("<ClinicalData StudyOID=\"demostudy30\" MetaDataVersionOID=\"1\">");
            // subject
            CdiscXml.Append("<SubjectData SubjectKey=\"");
            CdiscXml.Append(txtSubjectNumber.Text);
            CdiscXml.Append("\">");
            // site
            CdiscXml.Append("<SiteRef LocationOID=\"");
            CdiscXml.Append(txtSite.Text);
            CdiscXml.Append("\"/>");
            // 
            CdiscXml.Append("<StudyEventData StudyEventOID=\"\" StudyEventRepeatKey=\"1\">");
            // set up forced culture for use with dates
            IFormatProvider culture = new System.Globalization.CultureInfo("en-GB", true);
            // cover sheet
            if (chkCoverSheet.Checked)
            {
                // eform header
                CdiscXml.Append("<FormData FormOID=\"cover\" FormRepeatKey=\"1\">");
                CdiscXml.Append("<ItemGroupData ItemGroupOID=\"cover\" ItemGroupRepeatKey=\"0\">");
                // eform date
                CdiscXml.Append("<ItemData ItemOID=\"Covformdate\" Value=\"");
                CdiscXml.Append(dtEform.Value.ToString("dd/MM/yyyy", culture));
                CdiscXml.Append("\"/>");
                // subject no
                CdiscXml.Append("<ItemData ItemOID=\"subject\" Value=\"");
                CdiscXml.Append(txtSubjectNo.Text);
                CdiscXml.Append("\"/>");
                // subject initials
                CdiscXml.Append("<ItemData ItemOID=\"initials\" Value=\"");
                CdiscXml.Append(txtSubjectInitials.Text);
                CdiscXml.Append("\"/>");
                // dob
                CdiscXml.Append("<ItemData ItemOID=\"dobirth\" Value=\"");
                CdiscXml.Append(dtDOB.Value.ToString("dd/MM/yyyy", culture));
                CdiscXml.Append("\"/>");
                // sex
                if (rbFemale.Checked || rbMale.Checked)
                {
                    CdiscXml.Append("<ItemData ItemOID=\"sex\" Value=\"");
                    if (rbMale.Checked)
                    {
                        CdiscXml.Append("Male");
                    }
                    else
                    {
                        CdiscXml.Append("Female");
                    }
                    CdiscXml.Append("\"/>");
                }
                // eform footer
                CdiscXml.Append("</ItemGroupData>");
                CdiscXml.Append("</FormData>");
            }
            // visit date
            if (chkVisitDate.Checked)
            {
                // eform header
                CdiscXml.Append("<FormData FormOID=\"visitdate\" FormRepeatKey=\"1\">");
                CdiscXml.Append("<ItemGroupData ItemGroupOID=\"visitdate\" ItemGroupRepeatKey=\"0\">");
                // visit date
                CdiscXml.Append("<ItemData ItemOID=\"visitd\" Value=\"");
                CdiscXml.Append(dtVisit.Value.ToString("dd/MM/yyyy", culture));
                CdiscXml.Append("\"/>");
                // eform footer
                CdiscXml.Append("</ItemGroupData>");
                CdiscXml.Append("</FormData>");
            }
            // inclusion
            if (chkInclusion.Checked)
            {
                // eform header
                CdiscXml.Append("<FormData FormOID=\"eligibility\" FormRepeatKey=\"1\">");
                CdiscXml.Append("<ItemGroupData ItemGroupOID=\"eligibility\" ItemGroupRepeatKey=\"0\">");
                // healthy
                if (cbIsHealthy.SelectedIndex > -1)
                {
                    CdiscXml.Append("<ItemData ItemOID=\"healthy\" Value=\"");
                    CdiscXml.Append(cbIsHealthy.SelectedItem);
                    CdiscXml.Append("\"/>");
                }
                // blood pressure
                if (cbBloodPressure.SelectedIndex > -1)
                {
                    CdiscXml.Append("<ItemData ItemOID=\"bp\" Value=\"");
                    CdiscXml.Append(cbBloodPressure.SelectedItem);
                    CdiscXml.Append("\"/>");
                }
                // right age
                if (cbEighteen.SelectedIndex > -1)
                {
                    CdiscXml.Append("<ItemData ItemOID=\"rightage\" Value=\"");
                    CdiscXml.Append(cbEighteen.SelectedItem);
                    CdiscXml.Append("\"/>");
                }
                // body mass
                if (cbBodyMass.SelectedIndex > -1)
                {
                    CdiscXml.Append("<ItemData ItemOID=\"bmiok\" Value=\"");
                    CdiscXml.Append(cbBodyMass.SelectedItem);
                    CdiscXml.Append("\"/>");
                }
                // subject consent
                if (cbSigned.SelectedIndex > -1)
                {
                    CdiscXml.Append("<ItemData ItemOID=\"consent\" Value=\"");
                    CdiscXml.Append(cbSigned.SelectedItem);
                    CdiscXml.Append("\"/>");
                }
                // eform footer
                CdiscXml.Append("</ItemGroupData>");
                CdiscXml.Append("</FormData>");
            }
            // demographics
            if (chkDemographics.Checked)
            {
                // eform header
                CdiscXml.Append("<FormData FormOID=\"demography\" FormRepeatKey=\"1\">");
                CdiscXml.Append("<ItemGroupData ItemGroupOID=\"demography\" ItemGroupRepeatKey=\"0\">");
                // race
                if (cbRace.SelectedIndex > -1)
                {
                    CdiscXml.Append("<ItemData ItemOID=\"race\" Value=\"");
                    CdiscXml.Append(cbRace.SelectedIndex);
                    CdiscXml.Append("\"/>");
                }
                // subject status
                if (cbStatus.SelectedIndex > -1)
                {
                    CdiscXml.Append("<ItemData ItemOID=\"type\" Value=\"");
                    CdiscXml.Append(cbStatus.SelectedIndex);
                    CdiscXml.Append("\"/>");
                }
                // eform footer
                CdiscXml.Append("</ItemGroupData>");
                CdiscXml.Append("</FormData>");
            }
            // adverse event
            if (chkAdverseEvent.Checked)
            {
                // eform header
                CdiscXml.Append("<FormData FormOID=\"adverse\" FormRepeatKey=\"1\">");
                CdiscXml.Append("<ItemGroupData ItemGroupOID=\"adverse\" ItemGroupRepeatKey=\"0\">");
                // adverse event date
                CdiscXml.Append("<ItemData ItemOID=\"aedate\" Value=\"");
                CdiscXml.Append(dtEvent.Value.ToString("dd/MM/yyyy", culture));
                CdiscXml.Append("\"/>");
                // adverse event
                if (txtEvent.Text != "")
                {
                    CdiscXml.Append("<ItemData ItemOID=\"event\" Value=\"");
                    CdiscXml.Append(txtEvent.Text);
                    CdiscXml.Append("\">");
                    // connect to chosen eform date
                    CdiscXml.Append("<AuditRecord>");
                    CdiscXml.Append("<UserRef UserOID=\"\"/>");
                    CdiscXml.Append("<LocationRef LocationOID=\"\"/>");
                    CdiscXml.Append("<DateTimeStamp>");
                    CdiscXml.Append(dtEvent.Value.ToString("dd/MM/yyyy", culture));
                    CdiscXml.Append(" 00:00:00");
                    CdiscXml.Append("</DateTimeStamp>");
                    CdiscXml.Append("</AuditRecord>");
                    CdiscXml.Append("</ItemData>");
                }
                // serious
                if (cbSerious.SelectedIndex > -1)
                {
                    CdiscXml.Append("<ItemData ItemOID=\"serious\" Value=\"");
                    CdiscXml.Append(cbSerious.SelectedItem.ToString());
                    CdiscXml.Append("\">");
                    // connect to chosen eform date
                    CdiscXml.Append("<AuditRecord>");
                    CdiscXml.Append("<UserRef UserOID=\"\"/>");
                    CdiscXml.Append("<LocationRef LocationOID=\"\"/>");
                    CdiscXml.Append("<DateTimeStamp>");
                    CdiscXml.Append(dtEvent.Value.ToString("dd/MM/yyyy", culture));
                    CdiscXml.Append(" 00:00:00");
                    CdiscXml.Append("</DateTimeStamp>");
                    CdiscXml.Append("</AuditRecord>");
                    CdiscXml.Append("</ItemData>");
                }
                // eform footer
                CdiscXml.Append("</ItemGroupData>");
                CdiscXml.Append("</FormData>");
            }
            // close document
            CdiscXml.Append("</StudyEventData>");
            CdiscXml.Append("</SubjectData>");
            CdiscXml.Append("</ClinicalData>");
            CdiscXml.Append("</ODM>");
            // return
            return CdiscXml.ToString();
        }

        private bool ValidatePage()
        {
            // check have minimum of site & subject
            if (txtSite.Text == "" || txtSubjectNumber.Text == "")
            {
                MessageBox.Show("Please enter both 'Site' and 'Subject'");
                return false;
            }
            // and at least one section
            if (!(chkCoverSheet.Checked || chkVisitDate.Checked || chkInclusion.Checked
                    || chkDemographics.Checked || chkAdverseEvent.Checked))
            {
                MessageBox.Show("Please enter at least one eForm of data");
                return false;
            }
            return true;
        }

        private void btnResultXml_Click(object sender, EventArgs e)
        {
            // launch viewer
            if (_resultXml != "")
            {
                APIResults apiResults = new APIResults(_resultXml);
                apiResults.ShowDialog();
            }
            else
            {
                MessageBox.Show( "No Result XML to display!" );
            }
        }

        private void btnDataXml_Click(object sender, EventArgs e)
        {
            // launch viewer
            if (_dataXml != "")
            {
                APIResults apiResults = new APIResults(_dataXml);
                apiResults.ShowDialog();
            }
            else
            {
                MessageBox.Show("No Data XML to display!");
            }
        }

        private void ParseResultsXML(string resultXml)
        {
            StringReader tr = new StringReader(resultXml);

            XmlTextReader xmlTR = new XmlTextReader(tr);
            // retrieve buffersavestatus
            while (xmlTR.Read())
            {
                if (xmlTR.Name == "BufferMessageReport")
                {
                    string saveStatusResult = xmlTR.GetAttribute("BufferSaveStatus");
                    // <BufferMessageReport BufferSaveStatus="0">
                    int statusResult = Convert.ToInt16(saveStatusResult);
                    string resultString = "Unknown result!";
                    // deal with status result
                    switch (statusResult)
                    {
                        case 0:
                            {
                                resultString = "Stored response in buffer successfully";
                                break;
                            }
                        case 1:
                            {
                                resultString = "Message XML parse failure – message invalid";
                                break;
                            }
                        case 2:
                            {
                                resultString = "Clinical Trial Name invalid";
                                break;
                            }
                        case 3:
                            {
                                resultString = "Site invalid";
                                break;
                            }
                        case 4:
                            {
                                resultString = "Subject invalid";
                                break;
                            }
                        case 5:
                            {
                                resultString = "A data item code is invalid";
                                break;
                            }
                        case 6:
                            {
                                resultString = "A response value is invalid";
                                break;
                            }
                        case 7:
                            {
                                resultString = "The connection to the database has failed";
                                break;
                            }
                    }
                    txtLastResult.Text = resultString;
                    // found what we wanted - quit loop
                    break;
                }
            }
            // tidy up
            xmlTR.Close();
            tr.Close();
            tr.Dispose();
        }
    }
}