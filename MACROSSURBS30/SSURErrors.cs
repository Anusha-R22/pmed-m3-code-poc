/* ------------------------------------------------------------------------
 * File: SSURErrors.cs
 * Author: Nicky Johns
 * Copyright: InferMed, 2008, All Rights Reserved
 * Purpose: Error handling for Study/Site/User/Role import/export for API
 * Contains: Classes SSURErrors and SSURError
 * Date: March 2008
 * ------------------------------------------------------------------------
 * Revisions:
 * 
 * ------------------------------------------------------------------------*/

using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO;

namespace MACROSSURBS30
{
    /// <summary>
    /// Collection of Study/Site/User/Role association Error objects
    /// </summary>
    public class SSURErrors
    {
        /// <summary>
        /// Types of error
        /// </summary>
        public enum eSSURErr
        {
            None = 0,
            StudyNotExist = 1,
            SiteNotExist = 2,
            UserNotExist = 3,
            RoleNotExist = 4,
            StudySiteNotExist = 5,
            InvalidAction = 6,
            InvalidXML = 7
        }

        // Our list of errors
        List<SSURError> _errors = new List<SSURError>();

        /// <summary>
        /// Add an error to the collection
        /// </summary>
        /// <param name="errtype">Type of the error</param>
        /// <param name="assoc">The Association object which has the error</param>
        /// <param name="desc">Description of the error</param>
        public void Add(eSSURErr errtype, SSURAssoc assoc, string desc)
        {
            SSURError ssurerr = new SSURError(errtype, assoc, desc);

            _errors.Add(ssurerr);
        }

        /// <summary>
        /// Add an error of the given type and message
        /// </summary>
        /// <param name="errtype">Error type</param>
        /// <param name="study">Study name</param>
        /// <param name="site">Site code</param>
        /// <param name="user">User code</param>
        /// <param name="role">Role code</param>
        /// <param name="desc">Description of error</param>
        public void Add(eSSURErr errtype, string study, string site,
                    string user, string role, string desc)
        {
            SSURError ssurerr = new SSURError(errtype, study, site, user, role, desc);
            _errors.Add(ssurerr);
        }

        /// <summary>
        /// Return collection of errors as XML
        /// </summary>
        /// <returns>XML string</returns>
        public string ToXML()
        {
            // Nothing to do if no errors
            if (_errors.Count == 0) return "";

            // create xml in memory
            MemoryStream oMemStream = new MemoryStream();
            // create xmltextwriter - ASCII encoding
            XmlTextWriter tr = new XmlTextWriter(oMemStream, System.Text.Encoding.GetEncoding("ISO-8859-1"));

            tr.WriteStartDocument();
            tr.WriteStartElement("macroassocerrors");

            foreach (SSURError err in _errors)
            {
                // Output error details as XML
                err.AsXml(tr);
            }

            tr.WriteEndElement(); //macroassocerrors
            tr.WriteEndDocument(); //document
            tr.Flush();
            tr.Close();

            // collect xml string from memory stream
            ASCIIEncoding encoderAscii = new ASCIIEncoding();
            return encoderAscii.GetString(oMemStream.ToArray());
        }

        /// <summary>
        /// Study/Site/User/Role association Error object
        /// </summary>
        private class SSURError
        {
            // The attributes of the error
            SSURErrors.eSSURErr _errType = SSURErrors.eSSURErr.None;
            private SSURAssoc _assoc;
            private string _desc = "";

            /// <summary>
            /// Create a new error object
            /// </summary>
            /// <param name="errType">Error type</param>
            /// <param name="study">Study name, or "" if not known</param>
            /// <param name="site">Site code, or "" if not known</param>
            /// <param name="user">User name, or "" if not known</param>
            /// <param name="role">Role code, or "" if not known</param>
            /// <param name="desc">Text description of error</param>
            public SSURError(SSURErrors.eSSURErr errType, string study, string site,
                        string user, string role, string desc)
            {
                _errType = errType;
                _desc = desc;
                // Create the relevant SSUR object
                _assoc = new SSURAssoc(study, site, user, role);
            }

            /// <summary>
            /// Create a new error object
            /// </summary>
            /// <param name="errType">Error type</param>
            /// <param name="assoc">SSUR association object to which the error relates</param>
            /// <param name="desc">Text description of error</param>
            public SSURError(SSURErrors.eSSURErr errType, SSURAssoc assoc, string desc)
            {
                _errType = errType;
                _assoc = assoc;
                _desc = desc;
            }

            /// <summary>
            /// Write ourselves as a "macroassocerror" element to specified XML writer
            /// </summary>
            /// <param name="tr">XML writer for output</param>
            public void AsXml(XmlWriter tr)
            {
                tr.WriteStartElement("macroassocerror");

                tr.WriteAttributeString("msgtype", ((int)_errType).ToString());
                // Write the assoc attributes
                _assoc.WriteXMLAttributes(tr);
                tr.WriteAttributeString("msgdesc", _desc);

                tr.WriteEndElement();   //macroassocerror
            }
        }
    }
}
