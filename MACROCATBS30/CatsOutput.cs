/* ------------------------------------------------------------------------
 * File: CatsOutput.cs
 * Author: Nicky Johns
 * Copyright: InferMed, 2007, All Rights Reserved
 * Purpose: Category XML I/O for the MACRO 3.0 (and M4) API
 * Date: November 2007
 * ------------------------------------------------------------------------
 * Revisions:
 * NCJ 29 Nov 07 - Fixing bugs during WBT-ing
 * NCJ 5 Dec 07 - Sorted out updates of old/new categories
 * ------------------------------------------------------------------------*/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO;
using System.Data;
using MACROLOCKBS30;
using MACROAZRBBS30;

namespace MACROCATBS30
{
    public class CatsOutput
    {
        private string _dbCon = "";
        private string _userName = "";

        private CatErrors _catErrors;

        public CatsOutput(string dbCon, string userName)
        {
            _dbCon = dbCon;
            _userName = userName;
        }

        /// <summary>
        /// Get the XML representing all the categories specified in the request string
        /// </summary>
        /// <param name="xmlRequest">XLM request string</param>
        /// <returns>XML category output</returns>
        public string GetCatsXml(string xmlRequest)
        {
            XmlDocument doc = new XmlDocument();

             // Load the XML request string - bail out if invalid
            try { doc.LoadXml(xmlRequest); }
            catch { return ""; }

            // create xml in memory
            MemoryStream oMemStream = new MemoryStream();
            // create xmltextwriter - ASCII encoding
            //XmlTextWriter oXmlTw = new XmlTextWriter(oMemStream, Encoding.ASCII);
            XmlTextWriter tr = new XmlTextWriter(oMemStream, System.Text.Encoding.GetEncoding("ISO-8859-1"));

            tr.WriteStartDocument();

            tr.WriteStartElement("macrostudies");

            // Loop through study nodes
            foreach (XmlNode studyNode in doc.SelectNodes("//macrostudies/macrostudy"))
            {
                if (studyNode.Attributes["name"] != null)
                {
                    string study = studyNode.Attributes["name"].Value.ToString();
                    // Get the study ID
                    int studyId = CatSQL.GetStudyId(_dbCon, study);
                    if (studyId > 0)
                    {
                        // Write out this study's cat questions
                        tr.WriteStartElement("macrostudy");
                        tr.WriteAttributeString("name", study);
                        // Get all the requested categories for this study
                        WriteStudyCats(studyId, studyNode, tr);
                        tr.WriteEndElement();   // macrostudy
                    }
                }
            }

            tr.WriteEndElement(); //macrostudies
            tr.WriteEndDocument(); //document
            tr.Flush();
            tr.Close();

            // collect xml string from memory stream
            ASCIIEncoding encoderAscii = new ASCIIEncoding();
            return encoderAscii.GetString(oMemStream.ToArray());
            //return sb.ToString();
        }

        // Write out the requested category questions for this study
        private void WriteStudyCats(int studyId, XmlNode studyNode, XmlTextWriter tr)
        {
            tr.WriteStartElement("questions");
            XmlNodeList qNodeList = studyNode.SelectNodes("questions/question");
            // If no nodes, return ALL category questions
            if (qNodeList.Count == 0)
            {
                DataTable dt = CatSQL.GetCatQuestionCodes(_dbCon, studyId);
                foreach (DataRow row in dt.Rows) WriteQuestionCats(studyId, row[0].ToString(), tr);
            }
            else
                foreach (XmlNode qNode in qNodeList)
                {
                    if (qNode.Attributes["code"] != null) WriteQuestionCats(studyId, qNode.Attributes["code"].Value.ToString(), tr);
                }
            tr.WriteEndElement();   // questions
        }

        private void WriteQuestionCats(int studyId, string qcode, XmlTextWriter tr)
        {
            // Create a collection of the categories for this question
            // and load from the DB
            CatsM4 cats = new CatsM4(_dbCon, studyId, qcode, true);
            // Only do it if question exists (QuestionId > 0) and has categories
            if (cats.QuestionId > 0 && cats.Count > 0)
            {
                tr.WriteStartElement("question");
                tr.WriteAttributeString("code", qcode);
                // Get the categories as XML
                cats.AsXml(tr);
                tr.WriteEndElement();   // question
            }
        }

        /// <summary>
        /// Import category item updates
        /// </summary>
        /// <param name="XmlCats">XLM specification of category updates</param>
        /// <param name="errors">Error string - XML report of any errors or failed updates</param>
        /// <returns>Integer result code - 0 for Success (see API documentation for details)</returns>
        public int ImportCats(string XmlCats, out string errors)
        {
            errors = "";

            XmlDocument doc = new XmlDocument();
            try { doc.LoadXml(XmlCats); }
            // Return error code = 1 if couldn't read XML
            catch { return 1; }

            // Initialise our error collection
            _catErrors = new CatErrors();

            // Loop through study nodes
            foreach (XmlNode studyNode in doc.SelectNodes("//macrostudies/macrostudy"))
            {
                if (studyNode.Attributes["name"] != null)
                {
                    string study = studyNode.Attributes["name"].Value.ToString();
                    // Set study context in case of errors
                    _catErrors.Study = study;
                    // Get the study ID
                    int studyId = CatSQL.GetStudyId(_dbCon, study);
                    if (studyId > 0)
                    {
                        // Must lock the study during updates
                        string token = LockStudy(studyId);
                        if (token != "")
                        {
                            try
                            {
                                // If anything changes during import, rebuild AREZZO
                                if (ImportStudyCats(studyId, studyNode))
                                {
                                    // Call the VB MACROAZRBBS30 dll
                                    AzRebuildClass azRebuild = new AzRebuildClass();
                                    azRebuild.DoAREZZOUpdates(_dbCon, studyId);
                                }
                            }
                            finally
                            {
                                // Ensure locks not left behind
                                UnlockStudy(token, studyId);
                            }
                        }
                        else CreateError(CatErrors.eCatErr.StudyLocked, "Study in use");
                    }
                    else CreateError(CatErrors.eCatErr.StudyNotExist, "Study does not exist");
                }
                else CreateError(CatErrors.eCatErr.InvalidXML, "Missing study name");
            }
            // Now see what errors we had
            errors = _catErrors.ToXML();
            return (errors == "" ? 0 : 2);
        }

        // Import the category changes for this study
        // Returns True if anything changed and was saved
        private bool ImportStudyCats(int studyid, XmlNode studyNode)
        {
            int qid;
            int qtype;
            bool saved = false;
            // Get the questions
            XmlNodeList qNodeList = studyNode.SelectNodes("questions/question");
            foreach (XmlNode qNode in qNodeList)
            {
                if (qNode.Attributes["code"] != null)
                {
                    string qcode = qNode.Attributes["code"].Value.ToString();
                    // Set question context in case of errors
                    _catErrors.Question = qcode;
                    qid = CatSQL.GetDataItemId(_dbCon, studyid, qcode, out qtype);
                    if (qid == 0) CreateError(CatErrors.eCatErr.QuestionNotExist, "Question does not exist");
                    else
                    {
                        // Check question of type Category (data type = 1)
                        if (qtype != 1) CreateError(CatErrors.eCatErr.QuestionNotCat, "Question not of type category");
                        // If anything was saved update "saved" flag, otherwise leave it as it was
                        else saved = ImportQuestionCats(studyid, qcode, qNode) ? true : saved;
                    }
                }
                else CreateError(CatErrors.eCatErr.InvalidXML, "Missing question code");  // Invalid XML
            }
            // Was anything saved?
            return saved;
        }

        // Import all the categories for a particular question in a study
        // Return True if anything changed and was saved
        private bool ImportQuestionCats(int studyid, string qcode, XmlNode qNode)
        {
            // Load up the existing categories for this question
            CatsM4 cats = new CatsM4(_dbCon, studyid, qcode, true);
            foreach (XmlNode cNode in qNode.SelectNodes("categories/category"))
            {
                // Clear the way for new category
                _catErrors.CatCode = "";
                if (cNode.Attributes["code"] != null)
                {
                    string ccode = cNode.Attributes["code"].Value.ToString().Trim();
                    // Set catcode context for the error handler
                    _catErrors.CatCode = ccode;
                    // New or existing cat?
                    if (cats.CatExists(ccode)) ImportOldCat(cats, ccode, cNode);
                    else ImportNewCat(cats, ccode, cNode);
                }
                else CreateError(CatErrors.eCatErr.InvalidXML, "Missing category code");
            }
            // Get the sort order before saving
            if (qNode.Attributes["sort"] != null)
            {
                string sorter = qNode.Attributes["sort"].Value.ToString();
                try
                {
                    int esorter = Convert.ToInt32(sorter);
                    // Only set it if it's OK
                    if (cats.IsValidSort(esorter)) cats.SortOrder = (CatsM4.eCatSort)esorter;
                    else CreateError(CatErrors.eCatErr.InvalidXML, "Invalid sort value (" + sorter + ")");
                }
                catch
                {
                    CreateError(CatErrors.eCatErr.InvalidXML, "Invalid sort value (" + sorter + ")");
                }
            }
            // Only save if something's changed
            if (cats.NeedToSave())
            {
                cats.Save();
                // Clear cache entries so MACRO knows that study has changed
                DBLockClass locker = new DBLockClass();
                string nosite = "";
                int nosubject = -1;
                string notoken = "";
                locker.CacheInvalidate(ref _dbCon, ref studyid, ref nosite, ref nosubject, ref notoken);
                return true;
            }
            // Nothing happened
            return false;
        }

        // Import a new category item into question categories
        // from specified category node
        // ccode is new category code
        private void ImportNewCat(CatsM4 cats, string ccode, XmlNode cNode)
        {
            // Create new category
            CatM4 cat = new CatM4(ccode);
            // Did it work?
            if (ccode == "" || cat.Code != ccode)
            {
                // Invalid code - report error and stop
                CreateError(CatErrors.eCatErr.InvalidCode, "Invalid category code (" + ccode + ")");
                return;
            }
            if (cNode.Attributes["value"] == null)
            {
                // No Value - report error and stop
                CreateError(CatErrors.eCatErr.InvalidXML, "Missing category value");
                return;
            }
            string cvalue = cNode.Attributes["value"].Value.ToString().Trim();
            cat.Value = cvalue;
            // Did it work?
            if (cvalue == "" || cat.Value != cvalue)
            {
                CreateError(CatErrors.eCatErr.InvalidVal, "Invalid category value (" + cvalue + ")");
                return;
            }
            // Now deal with the Active value
            UpdateActive(cNode, cat);
            // And add it to the collection
            cats.AddCat(cat);
        }

        // Import changes to an existing category item into question categories
        // from specified category node
        // ccode is existing category code
        private void ImportOldCat( CatsM4 cats, string ccode, XmlNode cNode)
        {
            // Retrieve the existing category item
            CatM4 cat = cats.GetCat(ccode);
            if (cNode.Attributes["value"] != null)
            {
                string cvalue = cNode.Attributes["value"].Value.ToString().Trim();
                // Update the Value
                cat.Value = cvalue;
                // Did it work?
                if (cvalue == "" || cat.Value != cvalue) CreateError(CatErrors.eCatErr.InvalidVal, "Invalid category value (" + cvalue + ")");
            }
            // Finally deal with the Active value
            UpdateActive(cNode, cat);
        }

       
        private void CreateError(CatErrors.eCatErr errtype, string desc)
        {
            _catErrors.Add(errtype, desc);
        }


        // Get the "Active" value from a category node
        // and update the given category item appropriately
        // Note that there doesn't have to be an Active value
        private void UpdateActive(XmlNode cNode, CatM4 cat)
        {
            // If it's not there, do nothing
            if (cNode.Attributes["active"] == null) return;

            string val = cNode.Attributes["active"].Value.ToString().Trim().ToLower();
            switch (val)
            {
                case "true":
                    // Valid - update cat item
                    cat.Active = 1;
                    break;
                case "false":
                    // Valid - update cat item
                    cat.Active = 0;
                    break;
                default:
                    CreateError(CatErrors.eCatErr.InvalidActive, "Invalid Active value (" + val + ")");
                    break;
            }
        }

        /// <summary>
        /// Lock a study (using MACROLOCKBS30) and return unique token
        /// </summary>
        /// <param name="clinicalTrialId">Trial Id</param>
        /// <returns>Unique lock token, or "" if lock failed</returns>
        private string LockStudy(int clinicalTrialId)
        {
            DBLockClass locker = new DBLockClass();
            float wait = 0;
            string token;

            token = locker.LockStudy(_dbCon, _userName, clinicalTrialId, wait);
            switch (token)
            {
                case "0":
                case "1":
                case "2":
                    // These indicate that lock not obtained
                    token = "";
                    break;
                default:
                    break;
            }
            return (token);
        }


        /// <summary>
        /// Unlock macro study (using MACROLOCKBS30)
        /// </summary>
        /// <param name="token">Token assigned when study was locked</param>
        /// <param name="clinicalTrialId">Trial ID</param>
        private void UnlockStudy(string token, int clinicalTrialId)
        {
            DBLockClass locker = new DBLockClass();

            try
            {
                locker.UnlockStudy(_dbCon, token, clinicalTrialId);
            }
            catch
            {
            }
        }

    }
}
