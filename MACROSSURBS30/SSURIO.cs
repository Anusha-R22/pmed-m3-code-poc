/* ------------------------------------------------------------------------
 * File: SSURIO.cs
 * Author: Nicky Johns
 * Copyright: InferMed, 2008, All Rights Reserved
 * Purpose: Study/Site/User/Role IO using XML for MACRO 3.0 API
 * Contains: Class SSURIO
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
using System.Data;

namespace MACROSSURBS30
{
    /// <summary>
    /// Input and Output Study/Site/User/Role (SSUR) associations to/from MACRO 3.0 DB
    /// </summary>
    public class SSURIO
    {
        // Return values for Import Associations
        private const int IMPORT_SUCCESS = 0;
        private const int IMPORT_ERR = -1;
        private const int IMPORT_INVALIDXML = 1;
        private const int IMPORT_NOTALLDONE = 2;

        /// <summary>
        /// Export Study/Site/User/Role associations as XML from MACRO 3.0 DB
        /// </summary>
        /// <param name="xmlRequest">XML string requesting details</param>
        /// <param name="dbConn">MACRO database connection string</param>
        /// <returns>XML string containing required associations</returns>
        public string ExportAssocsXML(string xmlRequest, string dbConn)
        {
            XmlDocument doc = new XmlDocument();

            // Load the XML request string - bail out if invalid
            try { doc.LoadXml(xmlRequest); }
            catch { return ""; }

            // create xml in memory
            MemoryStream oMemStream = new MemoryStream();
            // create xmltextwriter - ASCII encoding
            XmlTextWriter tr = new XmlTextWriter(oMemStream, System.Text.Encoding.GetEncoding("ISO-8859-1"));

            tr.WriteStartDocument();
            tr.WriteStartElement("macroassociations");

            // Loop through study nodes
            foreach (XmlNode ssurNode in doc.SelectNodes("//macroassociations/macroassociation"))
            {
                // Collect all the attributes that are specified
                string study = "";
                if (ssurNode.Attributes["study"] != null)
                    study = ssurNode.Attributes["study"].Value.ToString();

                string site = "";
                if (ssurNode.Attributes["site"] != null)
                    site = ssurNode.Attributes["site"].Value.ToString();

                string user = "";
                if (ssurNode.Attributes["user"] != null)
                    user = ssurNode.Attributes["user"].Value.ToString();

                string role = "";
                if (ssurNode.Attributes["role"] != null)
                    role = ssurNode.Attributes["role"].Value.ToString();

                // Now we need to fetch the matching SSURs from the database
                SSURAssocs assocs = new SSURAssocs(dbConn, study, site, user, role);
                // Add them to the XML stream
                assocs.AssocsAsXML(tr);
            }

            tr.WriteEndElement(); //macroassociations
            tr.WriteEndDocument(); //document
            tr.Flush();
            tr.Close();

            // collect xml string from memory stream
            ASCIIEncoding encoderAscii = new ASCIIEncoding();
            return encoderAscii.GetString(oMemStream.ToArray());
        }

        /// <summary>
        /// Import Study/Site/User/Role associations from an XML string into MACRO 3.0 DB
        /// </summary>
        /// <param name="xmlAssocs">The XML associations</param>
        /// <param name="dbCon">MACRO database connection string</param>
        /// <param name="dbCode">MACRO database code</param>
        /// <param name="dbSecCon">Security database connection string</param>
        /// <param name="apiUserName">Name of current API user</param>
        /// <param name="xmlErrs">Errors, if any, as XML</param>
        /// <returns>0 if all OK, 1 if invalid XML, 2 if some errors (and xmlErrs contains details)</returns>
        public int ImportAssocsXML(string xmlAssocs, string dbCon, string dbCode,
                                string dbSecCon, string apiUserName, out string xmlErrs)
        {
            xmlErrs = "";
            int result = IMPORT_SUCCESS;

            XmlDocument doc = new XmlDocument();

            // Load the XML request string - bail out if invalid
            try { doc.LoadXml(xmlAssocs); }
            catch
            {
                xmlErrs = "XML string could not be read";
                return IMPORT_INVALIDXML;
            }

            // Initialise collection of errors
            SSURErrors errors = new SSURErrors();

            foreach (XmlNode ssurNode in doc.SelectNodes("//macroassociations/macroassociation"))
            {
                // Collect all the attributes that are specified
                string study = "";
                if (ssurNode.Attributes["study"] != null)
                    study = ssurNode.Attributes["study"].Value.ToString().Trim();

                string site = "";
                if (ssurNode.Attributes["site"] != null)
                    site = ssurNode.Attributes["site"].Value.ToString().Trim();

                string user = "";
                if (ssurNode.Attributes["user"] != null)
                    user = ssurNode.Attributes["user"].Value.ToString().Trim();

                string role = "";
                if (ssurNode.Attributes["role"] != null)
                    role = ssurNode.Attributes["role"].Value.ToString().Trim();

                string action = "";
                if (ssurNode.Attributes["action"] != null)
                    action = ssurNode.Attributes["action"].Value.ToString().Trim();

                // Store these initial (unchecked) values
                SSURAssoc assoc = new SSURAssoc(study, site, user, role);

                // Check everything is valid
                Boolean isValid = true;

                // Check we have a valid Action (if not, no point in checking the rest)
                if (isValid && action == "")
                {
                    errors.Add(SSURErrors.eSSURErr.InvalidXML, assoc, "Action attribute missing");
                    isValid = false;
                }
                // Check the Action
                if (isValid && !assoc.SetAction(action))
                {
                    errors.Add(SSURErrors.eSSURErr.InvalidAction, assoc, "Action '" + action + "' not recognised");
                    isValid = false;
                }
                // Check we have a Study
                if (isValid && study == "")
                {
                    errors.Add(SSURErrors.eSSURErr.InvalidXML, assoc, "Study attribute missing");
                    isValid = false;
                }
                // Check Study exists (and update with correct value from DB)
                if (isValid && !SSURSQL.StudyExists(dbCon, ref study))
                {
                    errors.Add(SSURErrors.eSSURErr.StudyNotExist, assoc, "Study not found");
                    isValid = false;
                }
                // Check we have a Site
                if (isValid && site == "")
                {
                    errors.Add(SSURErrors.eSSURErr.InvalidXML, assoc, "Site attribute missing");
                    isValid = false;
                }
                // Check Site exists
                if (isValid && !SSURSQL.SiteExists(dbCon, ref site))
                {
                    errors.Add(SSURErrors.eSSURErr.SiteNotExist, assoc, "Site not found");
                    isValid = false;
                }
                // Check Study/Site association exists
                if (isValid && !SSURSQL.StudySiteExists(dbCon, study, site))
                {
                    errors.Add(SSURErrors.eSSURErr.StudySiteNotExist, assoc, "Site is not participating in this study");
                    isValid = false;
                }
                // Check we have a User
                if (isValid && user == "")
                {
                    errors.Add(SSURErrors.eSSURErr.InvalidXML, assoc, "User attribute missing");
                    isValid = false;
                }
                // Check User exists (Security DB)
                if (isValid && !SSURSQL.UserExists(dbSecCon, ref user))
                {
                    errors.Add(SSURErrors.eSSURErr.UserNotExist, assoc, "User not found");
                    isValid = false;
                }
                // Check we have a Role
                if (isValid && role == "")
                {
                    errors.Add(SSURErrors.eSSURErr.InvalidXML, assoc, "Role attribute missing");
                    isValid = false;
                }
                // Check Role exists (Security DB)
                if (isValid && !SSURSQL.RoleExists(dbSecCon, ref role))
                {
                    errors.Add(SSURErrors.eSSURErr.RoleNotExist, assoc, "Role not found");
                    isValid = false;
                }
                // Something failed
                if (!isValid)
                    result = IMPORT_NOTALLDONE;
                else
                {
                    // Everything seems to be hunky dory
                    // Update with all the "official" (database) values
                    assoc.Update(study, site, user, role);
                    // Go ahead and save (add or delete) this association
                    assoc.SaveToDB(dbCon, dbSecCon, dbCode, apiUserName);
                }
            }

            // Collect up any errors
            xmlErrs = errors.ToXML();
            // Return overall result
            return result;
        }

    }
}

