/* ------------------------------------------------------------------------
 * File: TopCats.cs
 * Author: Nicky Johns
 * Copyright: InferMed, 2007, All Rights Reserved
 * Purpose: Top level class for Category processing in MACRO 3.0 API
 * Date: November 2007
 * ------------------------------------------------------------------------
 * Revisions:
 * 
 * ------------------------------------------------------------------------*/
using System;
using System.Collections.Generic;
using System.Text;

namespace MACROCATBS30
{
    /// <summary>
    /// Top level class for Category processing in MACRO 3.0 API
    /// </summary>
    public class TopCats
    {
        public TopCats() { }

        public bool ExportCats(string xmlRequest, string dbCon, string userName, out string xmlOut)
        {
            CatsOutput tabby = new CatsOutput(dbCon, userName);
            xmlOut = tabby.GetCatsXml(xmlRequest);
            return (xmlOut != "");
        }

        public int ImportCats(string xmlCats, string dbCon, string userName, out string xmlOut)
        {
            CatsOutput tabby = new CatsOutput(dbCon, userName);
            return tabby.ImportCats(xmlCats, out xmlOut);
        }
    }
}
