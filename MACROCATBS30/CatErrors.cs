/* ------------------------------------------------------------------------
 * File: CatM4.cs
 * Author: Nicky Johns
 * Copyright: InferMed, 2007, All Rights Reserved
 * Purpose: Category handling for the MACRO 3.0 (and M4) API
 * Contains: Classes CatErrors and CatError
 * Date: November 2007
 * ------------------------------------------------------------------------
 * Revisions:
 * 
 * ------------------------------------------------------------------------*/

using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO;

namespace MACROCATBS30
{
    /// <summary>
    /// Collection of errors which occur during category import
    /// </summary>
    public class CatErrors
    {
        public enum eCatErr
        {
            None = 0,
            StudyLocked = 1,
            StudyNotExist = 2,
            QuestionNotExist = 3,
            QuestionNotCat = 4,
            InvalidCode = 5,
            InvalidVal = 6,
            InvalidActive = 7,
            InvalidXML = 8
        }

        // The context of the error
        private string _study = "";
        public string Study
        {
            get { return _study; }
            // Clear Question and CatCode when setting Study
            set { _study = value; this.Question = ""; }
        }

        private string _question = "";
        public string Question
        {
            get { return _question; }
            // Clear CatCode when setting Question
            set { _question = value; _catCode = ""; }
        }

        private string _catCode = "";
        public string CatCode
        {
            get { return _catCode; }
            set { _catCode = value; }
        }

        // Our list of errors
        List<CatError> _errors = new List<CatError>();

        // Add an error of the given type and message
        // Assume that study, question etc. already set up
        public void Add(eCatErr errtype, string msg)
        {
            CatError ce = new CatError(errtype, _study, _question, _catCode, msg);
            _errors.Add(ce);
        }

        // Return our collection of errors as XML
        public string ToXML()
        {
            // Nothing to do if no errors
            if (_errors.Count == 0) return "";

            // create xml in memory
            MemoryStream oMemStream = new MemoryStream();
            // create xmltextwriter - ASCII encoding
            XmlTextWriter tr = new XmlTextWriter(oMemStream, System.Text.Encoding.GetEncoding("ISO-8859-1"));
            
            XmlDocument doc = new XmlDocument();

            tr.WriteStartDocument();
            tr.WriteStartElement("macrocaterrors");

            foreach (CatError ce in _errors)
            {
                // Output error details as XML
                ce.AsXml(tr);
            }

            tr.WriteEndElement(); //macrocaterrors
            tr.WriteEndDocument(); //document

            tr.Flush();
            tr.Close();

            // collect xml string from memory stream
            ASCIIEncoding encoderAscii = new ASCIIEncoding();
            return encoderAscii.GetString(oMemStream.ToArray());
        }
    }

    public class CatError
    {
        CatErrors.eCatErr _errtype = CatErrors.eCatErr.None;
        string _studyName = "";
        string _question = "";
        string _catcode = "";
        string _desc = "";

        // Any of these parameters can be "" if they're not relevant
        public CatError(CatErrors.eCatErr errtype, string study, string question,
                string catcode, string desc)
        {
            _errtype = errtype;
            _studyName = study;
            _question = question;
            _catcode = catcode;
            _desc = desc;
        }

        // Write ourselves to given XML writer
        public void AsXml(XmlWriter tr)
        {
            tr.WriteStartElement("macrocaterror");

            tr.WriteAttributeString("msgtype", ((int)_errtype).ToString());
            if (_studyName != "") tr.WriteAttributeString("study", _studyName);
            if (_question != "") tr.WriteAttributeString("question", _question);
            if (_catcode != "") tr.WriteAttributeString("categorycode", _catcode);
            tr.WriteAttributeString("msgdesc", _desc);

            tr.WriteEndElement();   //macrocaterror
        }
    }
}
