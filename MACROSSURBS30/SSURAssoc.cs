/* ------------------------------------------------------------------------
 * File: SSURAssoc.cs
 * Author: Nicky Johns
 * Copyright: InferMed, 2008, All Rights Reserved
 * Purpose: MACRO Study/Site/User/Role association object for MACRO 3.0 API
 * Date: March 2008
 * ------------------------------------------------------------------------
 * Revisions:
 * 
 * ------------------------------------------------------------------------*/

using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Xml;
using System.IO;

namespace MACROSSURBS30
{
    /// <summary>
    /// A collection of Study/Site/User/Role associations
    /// </summary>
    public class SSURAssocs
    {
        // List of association objects
        private List<SSURAssoc> _assocs = new List<SSURAssoc>();

        /// <summary>
        /// Get all the associations matching these values (any or all of which may be "")
        /// </summary>
        /// <param name="dbCon">MACRO database connection string</param>
        /// <param name="study">Study name (may be "")</param>
        /// <param name="site">Site code (may be "")</param>
        /// <param name="user">User name (may be "")</param>
        /// <param name="role">Role code (may be "")</param>
        public SSURAssocs(string dbCon, string study, string site, string user, string role)
        {
            _assocs = new List<SSURAssoc>();
            DataTable dt = SSURSQL.GetAssocs(dbCon, study, site, user, role);
            if (dt != null)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    // Create the SSUR association object from the DB info
                    SSURAssoc assoc = new SSURAssoc(dr["STUDYCODE"].ToString(),
                            dr["SITECODE"].ToString(), dr["USERNAME"].ToString(), dr["ROLECODE"].ToString());
                    // Add it to our collection
                    _assocs.Add(assoc);
                }
            }
        }

        /// <summary>
        /// Write out all the associations as XML elements
        /// Does NOT include the surrounding "macroassociations" tags
        /// </summary>
        /// <param name="tr">The XML writer for output</param>
        public void AssocsAsXML(XmlTextWriter tr)
        {
            foreach (SSURAssoc assoc in _assocs)
                assoc.WriteXML(tr);
        }

        /// <summary>
        /// Return this collection of associations as set of XML "macroassociations"
        /// </summary>
        /// <returns>XML string</returns>
        public string AsXML()
        {
            MemoryStream oMemStream = new MemoryStream();
            // create xmltextwriter - ASCII encoding
            XmlTextWriter tr = new XmlTextWriter(oMemStream, System.Text.Encoding.GetEncoding("ISO-8859-1"));

            tr.WriteStartDocument();
            tr.WriteStartElement("macroassociations");
            /// Write them all out
            AssocsAsXML(tr);
            tr.WriteEndElement(); //macroassociations
            tr.WriteEndDocument(); //document
            tr.Flush();
            tr.Close();

            // collect xml string from memory stream
            ASCIIEncoding encoderAscii = new ASCIIEncoding();
            return encoderAscii.GetString(oMemStream.ToArray());
        }
    }

    /// <summary>
    /// A Study/Site/User/Role assocation
    /// NB Do not save to DB unless all properties in correct case
    /// </summary>
    public class SSURAssoc
    {
        /// <summary>
        /// Whether to Add or Remove the association
        /// </summary>
        public enum eAction
        {
            Remove = 0,
            Add = 1,
            Nothing = 9
        }

        private string _study = "";
        public string Study
        {
            get { return _study; }
        }

        private string _site = "";
        public string Site
        {
            get { return _site; }
        }

        private string _user = "";
        public string User
        {
            get { return _user; }
        }

        private string _role = "";
        public string Role
        {
            get { return _role; }
        }

        private eAction _action = eAction.Nothing;
        public eAction Action
        {
            get { return _action; }
        }

        /// <summary>
        /// Set the action as a string, one of "add" or "remove".
        /// If neither of these, set to Nothing
        /// </summary>
        /// <param name="action">Action for this association</param>
        /// <returns>True if action was successfully set, False if invalid action</returns>
        public Boolean SetAction(string action)
        {
            switch (action)
            {
                case "add":
                    _action = eAction.Add;
                    return true;
                case "remove":
                    _action = eAction.Remove;
                    return true;
                default:
                    _action = eAction.Nothing;
                    return false;
            }
        }

        /// <summary>
        /// Create a new Association object for this Study, Site, User and Role
        /// </summary>
        /// <param name="study">Study name</param>
        /// <param name="site">Site code</param>
        /// <param name="user">User name</param>
        /// <param name="role">Role code</param>
        public SSURAssoc(string study, string site, string user, string role)
        {
            Update(study, site, user, role);
        }

        /// <summary>
        /// Check all our values are non-empty
        /// </summary>
        /// <returns>True if all SSUR properties are non-blank</returns>
        private Boolean AllNonEmpty()
        {
            return (_study != "" && _site != "" && _user != "" && _role != "");
        }

        /// <summary>
        /// Replace values with new ones (e.g. with values from DB with correct case)
        /// </summary>
        /// <param name="study">Study name</param>
        /// <param name="site">Site code</param>
        /// <param name="user">User name</param>
        /// <param name="role">Role code</param>
        public void Update(string study, string site, string user, string role)
        {
            _study = study;
            _site = site;
            _user = user;
            _role = role;
        }

        /// <summary>
        /// Does this association already exist in the DB?
        /// NB All properties must be in correct case
        /// </summary>
        /// <param name="dbCon">Database connection string</param>
        /// <returns>True if already exists; false otherwise</returns>
        private Boolean ExistsInDB(string dbCon)
        {
            if (AllNonEmpty())
                return SSURSQL.AssocExists(dbCon, _study, _site, _user, _role);
            // We have some blank entries
            return false;
        }

        /// <summary>
        /// Save this role association to the database.
        /// NB All properties must be in correct case to avoid corrupting MACRO DB!
        /// </summary>
        /// <param name="dbCon">MACRO database connection string</param>
        /// <param name="secDBCon">Security database connection string</param>
        /// <param name="dbCode">MACRO database code</param>
        /// <param name="apiUser">Current API user name</param>
        public void SaveToDB(string dbCon, string secDBCon, string dbCode, string apiUser)
        {
            // Check we have sensible values
            if (!AllNonEmpty())
                return;

            switch (_action)
            {
                case eAction.Add:
                    // Only do it if we haven't already got it
                    if (!ExistsInDB(dbCon))
                    {
                        // Save to the database
                        SSURSQL.InsertRole(dbCon, _study, _site, _user, _role);
                        // Update the UserDatabase table if necessary
                        SSURSQL.AddDBUser(secDBCon, dbCode, _user);
                        // Add the relevant messages to the MESSAGE table
                        SysMessages.SendNewUserRoleMessage(dbCon, secDBCon, apiUser, this);
                    }
                    break;
                case eAction.Remove:
                    // Only do it if we have already got it
                    if (ExistsInDB(dbCon))
                    {
                        // Delete from the database
                        SSURSQL.DeleteRole(dbCon, _study, _site, _user, _role);
                        // Add relevant message to the MESSAGE table
                        SysMessages.SendRemoveRoleMessage(dbCon, apiUser, this);
                        if (!SSURSQL.UserHasRolesInDB(dbCon, _user))
                        {
                            // Delete user/database link because user now has no roles in the DB
                            SSURSQL.DeleteDBUser(secDBCon, dbCode, _user);
                        }
                    }
                    break;
                default:
                    // Do nothing
                    break;
            }
        }

        /// <summary>
        /// Add object as XML "macroassociation" element to given XML writer
        /// </summary>
        /// <param name="tr">XML writer</param>
        public void WriteXML(XmlWriter tr)
        {
            tr.WriteStartElement("macroassociation");

            // Write the assoc attributes
            this.WriteXMLAttributes(tr);

            tr.WriteEndElement();   //macroassociation
        }

        /// <summary>
        /// Write out just the attributes of the object to the given XML writer
        /// Assume inside an XML element (does not write any element tags)
        /// </summary>
        /// <param name="tr">XML Writer</param>
        public void WriteXMLAttributes(XmlWriter tr)
        {
            if (_study != "") tr.WriteAttributeString("study", _study);
            if (_site != "") tr.WriteAttributeString("site", _site);
            if (_user != "") tr.WriteAttributeString("user", _user);
            if (_role != "") tr.WriteAttributeString("role", _role);
        }
    }
}
