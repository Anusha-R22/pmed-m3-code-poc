/* ------------------------------------------------------------------------
 * File: CatM4.cs
 * Author: Nicky Johns
 * Copyright: InferMed, 2007, All Rights Reserved
 * Purpose: Category handling for the MACRO 3.0 (and M4) API
 * Contains: Classes CatsM4 and CatM4
 * Date: November 2007
 * ------------------------------------------------------------------------
 * Revisions:
 * NCJ 29 Nov 07 - Fixing bugs during WBT-ing
 * NCJ 3 Dec 07 - Must update DataItemLength when saving Cats
 * NCJ 5 Dec 07 - Do updates under transaction control
 * ------------------------------------------------------------------------*/

using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Xml;
using System.Text.RegularExpressions;

namespace MACROCATBS30
{
    /// <summary>
    /// Collection of MACRO category items for a particular question
    /// </summary>
    public class CatsM4
    {
        // How should the categories be sorted?
        public enum eCatSort
        {
            None = 0,
            Code = 1,
            Value = 2
        }

        eCatSort _sortorder = eCatSort.None;

        // What's the sort order?
        public eCatSort SortOrder
        {
          get { return _sortorder; }
          set {
              eCatSort oldSort = _sortorder;
              _sortorder = value;
              if (_sortorder != oldSort) _dirty = true;
          }
        }

        // Does sort represent a valid Sort value?
        // Have to do it this way because eNums don't check their own values!
        public bool IsValidSort(int sort)
        {
            switch (sort)
            {
                case (int)eCatSort.None:
                case (int)eCatSort.Code:
                case (int)eCatSort.Value:
                    return true;
                default:
                    return false;
            }
        }

        // List of category items
        Dictionary<string, CatM4> _catList = new Dictionary<string, CatM4>();

        int _studyid = 0;
        //string _studyname = "";
        string _qcode;

        // The Question Id
        int _qid = 0;
        public int QuestionId
        {
            get { return _qid; }
        }

        string _dbCon = "";
        /// <summary>
        /// Database connection string
        /// </summary>
        public string DBCon
        {
            get { return _dbCon; }
            set { _dbCon = value; }
        }

        // The dirty flag
        bool _dirty = false;

        /// <summary>
        /// Create new empty Category collection for specified database, study and question
        /// </summary>
        /// <param name="dbcon">Database connection string</param>
        /// <param name="studyid">Study Id</param>
        /// <param name="qcode">Question Code</param>
        public CatsM4(string dbcon, int studyid, string qcode)
        {
            _studyid = studyid;
            _qcode = qcode;
            _dbCon = dbcon;
            // Get the dataitemID (may be 0 for non-existent question)
            _qid = CatSQL.GetDataItemId(dbcon, studyid, qcode);
        }

        /// <summary>
        /// Create new empty Category collection for specified database, study and question.
        /// If loaditems = true, categories are read from the DB
        /// </summary>
        /// <param name="dbcon">Database connection string</param>
        /// <param name="studyid">Study Id</param>
        /// <param name="qcode">Question Code</param>
        /// <param name="loaditems">True to immediately read the items from the DB, false to leave collection empty</param>
        public CatsM4(string dbcon, int studyid, string qcode, bool loaditems)
            : this(dbcon, studyid, qcode)
        {
            if (loaditems) this.Load();
        }

        /// <summary>
        /// Count of category items for this question
        /// </summary>
        public int Count
        {
            get { return _catList.Count; }
        }

        /// <summary>
        /// Add category item to collection. If order not filled in, an appropriate order value will be assigned
        /// </summary>
        /// <param name="cat">Category item</param>
        /// <returns>True if cat added successfully, false if not</returns>
        public bool AddCat( CatM4 cat )
        {
            string lcode = cat.Code.ToLower();
            // Check code doesn't already exist
            if (CodeExists(lcode)) return false;
            // Check code does not exist as Value either
            if (CodeExistsAsValue(lcode)) return false;

            // Set order if not already set or if order already exists
            if (cat.Order == 0 || OrderExists(cat.Order)) cat.Order = _catList.Count + 1;
            // Add it to collection
            _catList.Add(lcode, cat);
            // Set Dirty flag (to trigger a save)
            _dirty = true;
            return true;
        }

        // Does this category code already exist in our collection?
        public bool CatExists(string code)
        {
            return CodeExists(code.ToLower());
        }

        // Return category corresponding to this code
        public CatM4 GetCat(string code)
        {
            string lcode = code.ToLower();
            // Check code exists
            if (CodeExists(lcode)) return _catList[lcode];
            return null;
        }

        // Does this category Code already exist in our collection?
        // Assume lcode is lower case
        private bool CodeExists(string lcode)
        {
            return _catList.ContainsKey(lcode);
        }

        // Is there already a cat with this order value?
        private bool OrderExists(int order)
        {
            foreach (KeyValuePair<string, CatM4> kvp in _catList)
            {
                if (kvp.Value.Order == order) return true;
            }
            return false;
        }

        // Does this category Code already exist as a Value for a different code in our collection?
        // Assume lcode is lower case
        private bool CodeExistsAsValue(string lcode)
        {
            foreach (KeyValuePair<string,CatM4> kvp in _catList)
            {
                // NB Code and Value can match within a category item
                if (kvp.Key != lcode && kvp.Value.Value.ToLower() == lcode) return true;
            }
            return false;
        }

        private SortedDictionary<string, CatM4> SortCats(eCatSort sortsort)
        {
            // Transfer to a sorted dictionary
            SortedDictionary<string, CatM4> sortedcats = new SortedDictionary<string, CatM4>();
            string key;
            foreach (KeyValuePair<string, CatM4> kvp in _catList)
            {
                CatM4 cat = kvp.Value;
                // Decide what key to sort on
                switch (sortsort)
                {
                    case eCatSort.Code:
                        key =  cat.Code;
                        break;
                    case eCatSort.Value:
                        // Sort on Value but append Code to ensure unique keys in case of duplicate Values
                        key = cat.Value + cat.Code;
                        break;
                    default:
                        // Includes "None" - preserve existing ordering
                        key = cat.Order.ToString();
                        break;
               }
               sortedcats.Add(key, cat);
            }
            return sortedcats;
        }

        // Do we need to save?
        public bool NeedToSave()
        {
            // If our collection has been added to, we need to save
            if (_dirty == true) return true;
            // If we're going to sort, we need to save
            if (_sortorder != eCatSort.None) return true;
            // Now check each individual cat
            foreach (KeyValuePair<string, CatM4> kvp in _catList)
            {
                CatM4 cat = kvp.Value;
                // If anything's changed we need to save
                if (cat.Dirty == true) return true;
            }
            // No changes found
            return false;
        }

        // Save all the categories in our collection, sorting them first as necessary
        // NB Saving occurs even if nothing's changed - check NeedToSave first
        // Assumes study lock has been obtained!
        public void Save()
        {
            // Open a database connection for the complete process
            IDbConnection idbCon = CatSQL.GetConnection(_dbCon);
            // start transaction for whole delete-save cycle
            IDbTransaction dbTransaction = idbCon.BeginTransaction();
            try
            {
                // use transaction through command object
                IDbCommand dbComm = DataAccess.GetCommand(idbCon);
                dbComm.Transaction = dbTransaction;

                // First delete all the existing ones
                CatSQL.DelCats(dbComm, _studyid, _qid);

                // Now sort them before saving
                SortedDictionary<string, CatM4> sorted = SortCats(_sortorder);

                // We number the category rows numerically
                int id = 1;
                // Must calcualte the max. length as we go along
                int maxLength = 1;
                foreach (KeyValuePair<string, CatM4> kvp in sorted)
                {
                    CatM4 cat = kvp.Value;
                    // Ignore silly ones
                    if (cat.OKToSave())
                    {
                        // Set its unique id
                        cat.Order = id++;
                        cat.Save(dbComm, _studyid, _qid);
                        // Update max length
                        if (cat.Code.Length > maxLength) maxLength = cat.Code.Length;
                        if (cat.Value.Length > maxLength) maxLength = cat.Value.Length;
                    }
                }
                // Now finally update the data item length
                CatSQL.SaveDataItemLength(dbComm, _studyid, _qid, maxLength);
                // And commit the transaction
                dbTransaction.Commit();
            }
            catch (Exception ex)
            {
                // Don't commit anything if an error occurred
                dbTransaction.Rollback();
                throw ex;
            }
            // close database object
            if (idbCon.State == ConnectionState.Open)
            {
                // close
                idbCon.Close();
            }
            idbCon.Dispose();
        }

        // Load all the categories for the current question
        public void Load()
        {
            // Only do it if _qid > 0
            if (_qid == 0) return;
            CatM4 cat;
            DataTable dt = CatSQL.GetCats(_dbCon, _studyid, _qid);
            // Maybe found nothing
            if (dt == null) return;
            foreach (DataRow row in dt.Rows)
            {
                cat = new CatM4(row["ValueCode"].ToString(), row["ItemValue"].ToString(),
                                Convert.ToInt32(row["Active"].ToString()), Convert.ToInt32(row["ValueOrder"].ToString()));
                AddCat(cat);
            }
        }

        // Add the category collection to an XML stream
        public void AsXml(XmlWriter tr)
        {
            tr.WriteStartElement("categories");
            foreach (KeyValuePair<string, CatM4> kvp in _catList)
            {
                CatM4 cat = kvp.Value;
                tr.WriteStartElement("category");
                tr.WriteAttributeString("code", cat.Code);
                tr.WriteAttributeString("value", cat.Value);
                tr.WriteAttributeString("active", (cat.Active == 1 ? "true" : "false"));
                tr.WriteEndElement();   // category
            }
            tr.WriteEndElement();   // categories
        }

    }


    /// <summary>
    /// A MACRO Category item
    /// </summary>
    public class CatM4
    {
        private const int maxCodeLength = 15;
        private const int maxValueLength = 255;

        // Regular expressions for validating cat codes and values
        // Numeric - digits only
        private const string numeric = @"^[0-9][0-9]*$";
        // Alphanumeric - starts with letter, any chars (except space and invalid chars) follow
        private const string alphanumeric = @"^[A-Za-z_][^`""\|~\s]*$";
        // Category code can be numeric or alphanumeric
        private const string validcode = numeric + "|" + alphanumeric;
        // Category value can be anything except our invalid chars `"|~
        private const string validtext = @"^[^`""\|~]*$";

        // The Dirty flag - if true, cat needs saving
        bool _dirty = false;

        public bool Dirty
        {
            get { return _dirty; }
        }

        // Category code
        string _code = "";
        string _oldCode = "";
        // Do nothing if invalid Code passed in
        public string Code
        {
            get { return _code; }
            set {
                _code = (IsValidCode(value) ? value : _code);
                if (_code != _oldCode) _dirty = true;
            }
        }
        
        // Category value
        string _value = "";
        string _oldValue = "";
        // Do nothing if invalid Value passed in
        public string Value
        {
            get { return _value; }
            set {
                _value = (IsValidValue(value) ? value : _value);
                if (_value != _oldValue) _dirty = true;
            }
        }

        // Active = 0 (inactive) or 1 (active)
        int _active = 1;
        int _oldActive = 1;
        public int Active
        {
            get { return _active; }
            set
            {
                if (value == 0 || value == 1) _active = value;
                if (_active != _oldActive) _dirty = true;
            }
        }
        
        // Is this a valid category code?
        private bool IsValidCode(string catcode)
        {
            if (catcode == "") return false;
            if (catcode.Length > maxCodeLength) return false;
            // Category code can be numeric or alphanumeric
            Regex r = new Regex(validcode);
            return r.IsMatch(catcode);
        }

        // Is this a valid category value?
        private bool IsValidValue(string catval)
        {
            if (catval == "") return false;
            if (catval.Length > maxValueLength) return false;
            // Category value can be any valid text
            Regex r = new Regex(validtext);
            return r.IsMatch(catval);
        }

        // OK to save if _code and _value are OK
        public bool OKToSave()
        {
            return (_code != "" && _value != "" && _order > 0); 
        }


        // Category order (also doubles as ValueID)
        int _order = 0;
        public int Order
        {
            get { return _order; }
            set { _order = value; }
        }

        public CatM4(string code)
        {
            // Use public property which triggers the value checks
            this.Code = code;
        }

        // Instantiate an existing category item (dirty flag is NOT set)
        public CatM4(string code, string value, int active)
        {
            // Use private properties which don't trigger the value checks
            _code = code;
            _oldCode = code;
            _value = value;
            _oldValue = value;
            _active = active;
            _oldActive = active;
        }

        /// <summary>
        /// Create a new category item WITHOUT validating the parameters (assumed valid!)
        /// </summary>
        /// <param name="code">Category code</param>
        /// <param name="value">Category value</param>
        /// <param name="active">Active (0 or 1)</param>
        /// <param name="order">Category order</param>
        public CatM4(string code, string value, int active, int order)
            : this(code, value, active)
        {
            _order = order;
            _dirty = false;
        }

        /// <summary>
        /// Save this category to the database. Does nothing if category not valid.
        /// </summary>
        /// <param name="dbComm">Open Database command object</param>
        /// <param name="studyid">Study Id</param>
        /// <param name="dataitemid">Data Item ID</param>
        public void Save(IDbCommand dbComm, int studyid, int dataitemid)
        {
            // Don't save an invalid cat
            if (OKToSave()) CatSQL.SaveCat(dbComm, studyid, dataitemid, _code, _value, _active, _order);
        }


    }
}
