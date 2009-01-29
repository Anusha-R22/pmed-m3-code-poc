using System;
using System.Data;
using InferMed.MACRO.ClinicalCoding.Plugins;
using InferMed.MACRO.ClinicalCoding.MACROCCBS30;
using InferMed.Components;
using System.Collections;
using System.IO;
using System.Xml;

namespace InferMed.MACRO.ClinicalCoding.MACRO_CC
{
	/// <summary>
	/// AutoEncoding functions
	/// </summary>
	public class MedDRA
	{
		//callback delegate
		public delegate void DShowProgress( string prog, bool msg );
		//callback delegate event handler
		public static event DShowProgress DShowProgressEvent;
		
		//callback delegate
		public delegate void DAddTermToList( AutoCodedTermHistory t );
		//callback delegate event handler
		public static event DAddTermToList DAddTermToListEvent;


		public enum CCDBFieldType
		{
			Integer = 4, VarChar = 10
		}

		private const string _LLT_TABLE = "t1_low_level_term";
		private const string _PT_TABLE = "t1_pref_term";
		private const string _HLT_TABLE = "t1_hlt_pref_term";
		private const string _HLGT_TABLE = "t1_hlgt_pref_term";
		private const string _SOC_TABLE = "t1_soc_term";
		private const string _DICTIONARY_INFO_TABLE = "DICTIONARY_INFO";

		private const string _DEF_USER_SETTINGS_FILE = @"C:\Program Files\InferMed\MACRO 3.0\MACROUserSettings30.txt";
		private const string _DEF_MACRO_PATH = @"C:\Program Files\InferMed\MACRO 3.0\";
		private const string _SETTINGS_FILE = @"MACROSettings30.txt";
		private const string _USER_SETTINGS_FILE = "settingsfile";
		private const string _CC_DB_CONNECTION = "ccsecurityPath";

		private MedDRA()
		{
		}

		/// <summary>
		/// Returns the MACRO usersettings file path from the settings file
		/// </summary>
		/// <returns>User settings file path</returns>
		private static string GetUserSettingsFilePath()
		{
			string settingsFile = AppDomain.CurrentDomain.BaseDirectory  + _SETTINGS_FILE;
			if( !File.Exists( settingsFile ) )
			{
				Exception ex = new Exception( "Settings file '" + settingsFile + "' not found" );
				throw ex;
			}
			IMEDSettings20 iset = new IMEDSettings20( settingsFile );
			string userSettingsFile = iset.GetKeyValue( _USER_SETTINGS_FILE, _DEF_USER_SETTINGS_FILE );
			if( !File.Exists( userSettingsFile ) )
			{
				Exception ex = new Exception( "User settings file '" + userSettingsFile + "' not found" );
				throw ex;
			}
			return( userSettingsFile );
		}

		/// <summary>
		/// Returns a setting from the MACRO 3.0 user settings file
		/// </summary>
		/// <param name="key">Setting key name</param>
		/// <param name="defaultVal">Default value to be returned if the key is not found</param>
		/// <returns>MACRO setting value</returns>
		public static string GetSetting(string key, string defaultVal)
		{
			IMEDSettings20 iset = new IMEDSettings20( GetUserSettingsFilePath() );
			return( iset.GetKeyValue( key, defaultVal ) );
		}

		/// <summary>
		/// Updates a setting in the MACRO 3.0 user settings file
		/// </summary>
		/// <param name="key">Setting key name</param>
		/// <param name="val">Setting value</param>
		public static void SetSetting(string key, string val)
		{
			string settingsFile = GetUserSettingsFilePath();
			IMEDSettings20 iset = new IMEDSettings20( settingsFile );
			iset.SetKeyValue( key, val );
		}

		/// <summary>
		/// AutoEncode MedDRA terms
		/// </summary>
		/// <param name="con"></param>
		/// <param name="userName"></param>
		/// <param name="userNameFull"></param>
		/// <param name="clinicalTrialId"></param>
		/// <param name="trialSite"></param>
		public static void AutoEncode(string con, string userName, string userNameFull, int clinicalTrialId, 
			ArrayList trialSite)
		{
			Plugins.MedDRA m = null;
			AutoCodedTermHistory t = null;
			Dictionaries d = new Dictionaries();
			int recs = 0, currentRec = 0;
		
			//load all dictionaries
			d.Init( con );
			//if any dictionaries are installed
			if( d.Count > 0 )
			{
				//get a dataset of all responses to be auto-encoded
				DShowProgressEvent( "Getting uncoded responses", false );
				DataSet ds = GetUncodedResponses( con, clinicalTrialId, trialSite );
				//if there are any responses to auto-encode
				if( ds.Tables[0].Rows.Count > 0 )
				{
					recs = ds.Tables[0].Rows.Count;
					currentRec = 1;
					//get the connection string for the first dictionary in the collection
					//the server parameters will be the same for all dictionaries
					Dictionary tempD  = ( Dictionary )d.DictionaryList[0];
					string custom = tempD.Custom;
					//part initialise the meddra plugin
					DShowProgressEvent( "Connecting to server", false );
					m = new Plugins.MedDRA( custom );

					if( m.Status == CCDataServer.ServerStatus.Connected )
					{
						//loop through the dataset
						foreach( DataRow row in ds.Tables[0].Rows )
						{
							int totalMatches = 0;
							//get the dictionary for this response
							tempD = d.DictionaryFromId( System.Convert.ToInt32( row["DICTIONARYID"].ToString() ) );
							//set the dictionary in the encoder
							DShowProgressEvent( "Loading dictionary, " + tempD.Name + " " + tempD.Version + " (term " 
								+ currentRec + " of " + recs + ")", false );
							m.SetDictionary( tempD.Custom );

							if( m.Status == CCDataServer.ServerStatus.Ready )
							{
								//autocode the value
								DShowProgressEvent( "AutoEncoding " + row["PERSONID"].ToString() + ", " + row["RESPONSETASKID"].ToString() 
									+ " [" + row["REPEATNUMBER"].ToString() + "]" + " (term " + currentRec + " of " + recs + ")", false );
								string codedValue = m.AutoEncode( row["RESPONSEVALUE"].ToString(), ref totalMatches );
					
								t = new AutoCodedTermHistory();
								t.Matches = totalMatches;
								if( codedValue != "")
								{
									//exact match found
									DShowProgressEvent( "Loading term history " + row["PERSONID"].ToString() + ", " 
										+ row["RESPONSETASKID"].ToString() + " [" + row["REPEATNUMBER"].ToString() + "]" 
										+ " (term " + currentRec + " of " + recs + ")", false );
									t.InitAuto( con, clinicalTrialId, row["TRIALSITE"].ToString(),
										System.Convert.ToInt32( row["PERSONID"].ToString() ), 
										System.Convert.ToInt32( row["VISITID"].ToString() ), 
										System.Convert.ToInt16( row["VISITCYCLENUMBER"].ToString() ), 
										System.Convert.ToInt32( row["CRFPAGEID"].ToString() ), 
										System.Convert.ToInt16( row["CRFPAGECYCLENUMBER"].ToString() ), 
										System.Convert.ToInt32( row["RESPONSETASKID"].ToString() ),
										System.Convert.ToInt16( row["REPEATNUMBER"].ToString() ), tempD );

									string rfc = ( t.RequiresRFC ) ? "Automatically encoded" : "";
									t.SetCode( tempD.Name, tempD.Version, codedValue, userName, 
										userNameFull, row["RESPONSEVALUE"].ToString(),
										System.Convert.ToDouble( row["RESPONSETIMESTAMP"].ToString() ), 
										System.Convert.ToInt16( row["RESPONSETIMESTAMP_TZ"].ToString() ),
										rfc, true );
							
									DAddTermToListEvent( t );	
								}
								else
								{
									//exact match not found
									DShowProgressEvent( "Loading empty term history " + row["PERSONID"].ToString() + ", " 
										+ row["RESPONSETASKID"].ToString() + " [" + row["REPEATNUMBER"].ToString() + "]" 
										+ " (term " + currentRec + " of " + recs + ")", false );
									t.InitEmpty( clinicalTrialId, row["TRIALSITE"].ToString(), 
										System.Convert.ToInt32( row["PERSONID"].ToString() ),
										System.Convert.ToInt32( row["VISITID"].ToString() ), 
										System.Convert.ToInt16( row["VISITCYCLENUMBER"].ToString() ), 
										System.Convert.ToInt32( row["CRFPAGEID"].ToString() ), 
										System.Convert.ToInt16( row["CRFPAGECYCLENUMBER"].ToString() ), 
										System.Convert.ToInt32( row["RESPONSETASKID"].ToString() ),
										System.Convert.ToInt16( row["REPEATNUMBER"].ToString() ),
										row["RESPONSEVALUE"].ToString(),
										System.Convert.ToDouble( row["RESPONSETIMESTAMP"].ToString() ), 
										System.Convert.ToInt16( row["RESPONSETIMESTAMP_TZ"].ToString() ), tempD );

									DAddTermToListEvent( t );
								}
							}
							else
							{
								//cant load dictionary
								DShowProgressEvent( "Unable to load dictionary " + tempD.Name + " " + tempD.Version, true );
							}
							currentRec++;
						}
					}
					else
					{
						//cant connect to server
						DShowProgressEvent( "Unable to connect to server, autoencode terminated", true );
					}
					DShowProgressEvent( "Disconnecting from server", false );
					m.Dispose();
				}
			}
		}

		/// <summary>
		/// Get responses that require autoencoding
		/// </summary>
		/// <param name="con"></param>
		/// <param name="clinicalTrialId"></param>
		/// <param name="trialSite"></param>
		/// <returns></returns>
		private static DataSet GetUncodedResponses( string con, long clinicalTrialId, ArrayList trialSite  )
		{
			string sql = "";

			sql = "SELECT DATAITEMRESPONSE.TRIALSITE, DATAITEMRESPONSE.PERSONID, DATAITEMRESPONSE.VISITID, DATAITEMRESPONSE.VISITCYCLENUMBER, DATAITEMRESPONSE.CRFPAGEID, "
				+ "DATAITEMRESPONSE.CRFPAGECYCLENUMBER, DATAITEMRESPONSE.RESPONSETASKID, DATAITEMRESPONSE.REPEATNUMBER, "
				+ "DATAITEMRESPONSE.RESPONSEVALUE, DATAITEMRESPONSE.RESPONSETIMESTAMP, DATAITEMRESPONSE.RESPONSETIMESTAMP_TZ, "
				+ "DATAITEM.DICTIONARYID "
				+ "FROM DATAITEMRESPONSE, DATAITEM "
				+ "WHERE DATAITEMRESPONSE.CLINICALTRIALID = DATAITEM.CLINICALTRIALID "
				+ "AND DATAITEMRESPONSE.DATAITEMID = DATAITEM.DATAITEMID "
				+ "AND DATAITEMRESPONSE.CLINICALTRIALID = " + clinicalTrialId + " "
				+ "AND TRIALSITE " + GetSiteSQL( trialSite )
				+ "AND DATAITEMRESPONSE.CODINGSTATUS = 1 AND DATAITEMRESPONSE.RESPONSEVALUE <> ''";

			return( DataAccess.GetDataSet( con, sql ) );
		}

		/// <summary>
		/// Get site SQL
		/// </summary>
		/// <param name="trialSite"></param>
		/// <returns></returns>
		private static string GetSiteSQL( ArrayList trialSite )
		{
			string sql = "";

			if( trialSite.Count == 1 )
			{
				sql = "= '" + trialSite[0] + "' ";
			}
			else
			{
				sql = "IN (";
				for( int n = 0; n < trialSite.Count; n++ )
				{
					sql += "'" + trialSite[n] + "', ";
				}
				sql = sql.Substring(0, ( sql.Length - 2 ) );
				sql +=") ";
			}
			return( sql );
		}

		/// <summary>
		/// Get cc database connection string
		/// </summary>
		/// <returns></returns>
		public static string GetDBConnectionString()
		{
			string dbCon = "";

			dbCon = GetSetting( _CC_DB_CONNECTION, "" );
			if( dbCon == "" )
			{
				DBConnectionForm f = new DBConnectionForm();
				f.ShowDialog();
				dbCon = f.DBCon;
				f.Dispose();

				if( dbCon != "" )
				{
					string eDbCon = IMEDEncryption.EncryptString( dbCon );
					SetSetting( _CC_DB_CONNECTION, eDbCon );
				}
			}
			else
			{
				dbCon = IMEDEncryption.DecryptString( dbCon );
			}

			return( dbCon );
		}

		/// <summary>
		/// Create cc table
		/// </summary>
		/// <param name="dbCon"></param>
		/// <param name="tableName"></param>
		/// <param name="xmlNlFields"></param>
		/// <param name="drop"></param>
		private static void CreateDicTable( IDbConnection dbCon, string tableName, XmlNodeList xmlNlFields, bool drop )
		{
			//get db type
			IMEDDataAccess.ConnectionType cType = IMEDDataAccess.CalculateConnectionType( dbCon );

			//drop table
			if( drop ) DropDicTable( dbCon, tableName );
			
			//build the create table sql
			string sql = "CREATE TABLE " + tableName + " (";

			foreach( XmlNode xmlNField in xmlNlFields )
			{
				sql += GetColumnName( cType, xmlNField.Attributes["NAME"].Value.ToString() ) + " "
					+ GetColumnType( cType, System.Convert.ToInt32( xmlNField.Attributes["TYPE"].Value.ToString() ),
					xmlNField.Attributes["LENGTH"].Value.ToString() ) + ", ";
			}
			sql = sql.Substring( 0, ( sql.Length - 2 ) ) + ")";

			DataAccess.RunSQL( dbCon, sql );
		}

		/// <summary>
		/// Get table column name
		/// </summary>
		/// <param name="cType"></param>
		/// <param name="colName"></param>
		/// <returns></returns>
		private static string GetColumnName( IMEDDataAccess.ConnectionType cType, string colName )
		{
			string name = "";

			switch( cType )
			{
				case IMEDDataAccess.ConnectionType.SQLServer:
					name = "[" + colName + "]";
					break;
				case IMEDDataAccess.ConnectionType.Oracle:
					//name = "\"" + colName + "\"";
					name = colName;
					break;
			}
			return( name );
		}

		/// <summary>
		/// Get table column type
		/// </summary>
		/// <param name="cType"></param>
		/// <param name="DBFieldType"></param>
		/// <param name="DBFieldLength"></param>
		/// <returns></returns>
		private static string GetColumnType( IMEDDataAccess.ConnectionType cType, int DBFieldType, string DBFieldLength )
		{
			string colType = "";

			switch( cType )
			{
				case IMEDDataAccess.ConnectionType.SQLServer:
				switch( DBFieldType )
				{
					case ( int )CCDBFieldType.Integer:
						colType = "int";
						break;
					case ( int )CCDBFieldType.VarChar:
						colType = "VarChar(" + DBFieldLength + ")";
						break;
				}
					break;
				case IMEDDataAccess.ConnectionType.Oracle:
				switch( DBFieldType )
				{
					case ( int )CCDBFieldType.Integer:
						colType = "int";
						break;
					case ( int )CCDBFieldType.VarChar:
						colType = "VarChar2(" + DBFieldLength + ")";
						break;
				}
					break;
			}
			return( colType );
		}

		/// <summary>
		/// Get a value safe for including in an sql string
		/// </summary>
		/// <param name="val"></param>
		/// <param name="DBFieldType"></param>
		/// <returns></returns>
		private static string GetSqlValue( string val, int DBFieldType )
		{
			string sqlValue = "";

			val = val.Trim();
			switch( DBFieldType )
			{
				case ( int )CCDBFieldType.Integer:
					sqlValue = ( val == "" ) ? "0" : val;
					break;
				case ( int )CCDBFieldType.VarChar:
					sqlValue = "'" + DataAccess.SQLTerm( val ) + "'";
					break;
			}
			return( sqlValue );
		}

		/// <summary>
		/// Drop a cc table
		/// </summary>
		/// <param name="dbCon"></param>
		/// <param name="tableName"></param>
		private static void DropDicTable( IDbConnection dbCon, string tableName )
		{
			try
			{
				//drop table
				string sql = "DROP TABLE " + tableName;
				DataAccess.RunSQL( dbCon, sql );
			}
			catch( Exception ex )
			{
				//ignore errors
			}
		}

		/// <summary>
		/// Import a cc dictionary
		/// </summary>
		/// <param name="con"></param>
		/// <param name="mapPath"></param>
		/// <param name="importDir"></param>
		/// <param name="dName"></param>
		/// <param name="dVersion"></param>
		/// <param name="stlDName"></param>
		/// <param name="prefix"></param>
		/// <param name="pluginNamespace"></param>
		public static void Import( string con, string mapPath, string importDir, out string dName, out string dVersion,
			out string stlDName, out string prefix, out string pluginNamespace )
		{
			System.Xml.XmlDocument x = new XmlDocument();
			StreamReader sr = null;
			IMEDDataAccess imedData = new IMEDDataAccess();
			IDbConnection dbCon = null;
			ArrayList createdTables = new ArrayList();
			dName = ""; dVersion = ""; stlDName = ""; prefix = "";
			string sql = "";
			int lltCount = 0, ptCount = 0, hltCount = 0, hlgtCount = 0, socCount = 0, bigCount;


			try
			{
				// open ImedDataAccess class
				IMEDDataAccess.ConnectionType imedConnType;
				// calculate connection type
				imedConnType=IMEDDataAccess.CalculateConnectionType( con );
				// get connection & open
				dbCon = imedData.GetConnection( imedConnType, con );
				dbCon.Open();

				x.Load( mapPath );
				dName = x.SelectSingleNode( "//MAP" ).Attributes["NAME"].Value.ToString();
				dVersion = x.SelectSingleNode( "//MAP" ).Attributes["VERSION"].Value.ToString();
				stlDName = x.SelectSingleNode( "//MAP" ).Attributes["STLVERSION"].Value.ToString();
				prefix = x.SelectSingleNode( "//MAP" ).Attributes["TABLEPREFIX"].Value.ToString();
				pluginNamespace = x.SelectSingleNode( "//MAP" ).Attributes["NAMESPACE"].Value.ToString();

				//loop through each meddra file
				foreach( XmlNode xmlNFile in x.SelectNodes( "//MAP/FILE" ) )
				{
					//get the path and name of file
					string meddraFile = xmlNFile.Attributes["NAME"].Value.ToString();
					string meddraPath = importDir + @"\" + meddraFile;
					//get the delimiter used in the file
					string delimiter = xmlNFile.Attributes["DELIMITER"].Value.ToString();
					//zero row count
					int rowCount = 0;

					//if the file exists
					if( File.Exists( meddraPath ) )
					{
						string tableName = prefix + xmlNFile.Attributes["TABLE"].Value.ToString();
						sql = "";

						//get the field maps
						XmlNodeList xmlNlFields = x.SelectNodes( "//MAP/TABLE[@NAME='" + xmlNFile.Attributes["TABLE"].Value.ToString() 
							+ "']/FIELD" );
	
						//check if this table has been created already
						if( !IsInArray( createdTables, tableName ) )
						{
							//create the table
							DShowProgressEvent( "file " + meddraFile + ", create " + tableName, false );
							CreateDicTable( dbCon, tableName, xmlNlFields, true );
							createdTables.Add( tableName );
						}

						//build a list of columns for the insert sql
						string cols = "(";
						foreach( XmlNode xmlNField in xmlNlFields )
						{
							cols += GetColumnName( imedConnType, xmlNField.Attributes["NAME"].Value.ToString() ) + ", ";
						}
						cols = cols.Substring(0, ( cols.Length - 2 ) ) + ")";
						
						//open the dictionary delimited file
						DShowProgressEvent( "File " + meddraFile + ", opening", false );
						sr = File.OpenText( meddraPath );

						string input;
						char [] del = delimiter.ToCharArray();
						//read each line of the file in
						while ( ( input = sr.ReadLine() ) != null )
						{
							//create insert sql
							sql = "INSERT INTO " + tableName + " " + cols;

							//loop through all the fields expected in each row adding them as values
							string[] meddraParams = input.Split( del );
							string vals = "VALUES (";
							foreach( XmlNode xmlNField in xmlNlFields )
							{
								vals += GetSqlValue( meddraParams[( System.Convert.ToInt32( xmlNField.Attributes["SEQUENCE"].Value.ToString() ) - 1 )],
									System.Convert.ToInt32( xmlNField.Attributes["TYPE"].Value.ToString() ) ) + ", ";
							}
							vals = vals.Substring(0, ( vals.Length - 2 ) ) + ")";
							sql += " " + vals;
							DShowProgressEvent( "file " + meddraFile + ", insert " + tableName + " " + vals, false );
							rowCount++;
							DataAccess.RunSQL( dbCon, sql );
						}

						sr.Close();
					}
					else
					{
						Exception ex = new Exception( "MedDRA file missing : " + meddraFile );
						throw ex;
					}

					switch( xmlNFile.Attributes["TABLE"].Value.ToString() )
					{
						case _LLT_TABLE:
							lltCount += rowCount;
							break;
						case _PT_TABLE:
							ptCount += rowCount;
							break;
						case _HLT_TABLE:
							hltCount += rowCount;
							break;
						case _HLGT_TABLE:
							hlgtCount += rowCount;
							break;
						case _SOC_TABLE:
							socCount += rowCount;
							break;
					}
				}

				//build large search table
				BuildBigSearchTable( dbCon, x, prefix, out bigCount );

				//update dictionary info table
				sql = "INSERT INTO DICTIONARY_INFO (DICT_TYPE, DICT_VERSION, DICT_LANGUAGE, DICT_PREFIX, FLAT_COUNT, COUNT01, "
					+ "COUNT02, COUNT03, COUNT04, COUNT05, ALLOWED_GROUP_1) "
					+ "VALUES ('MedDRA', '" + stlDName + "', 'ENGLISH', '" + prefix + "', " + bigCount + ", " + lltCount + ", " + ptCount 
					+ ", " + hltCount + ", " + hlgtCount + ", " + socCount + ", 'USER')";
				DataAccess.RunSQL( dbCon, sql );
			}
			finally
			{
				dbCon.Close();
				dbCon.Dispose();
			}
		}

		/// <summary>
		/// Create big search table
		/// </summary>
		/// <param name="dbCon"></param>
		/// <param name="x"></param>
		/// <param name="prefix"></param>
		/// <param name="bigCount"></param>
		private static void BuildBigSearchTable( IDbConnection dbCon, XmlDocument x, string prefix, out int bigCount )
		{
			string sql = "", cols = "", tblName = "";
			DataSet ds = null;

			try
			{
				tblName = prefix + "BIG_SEARCH_TABLE";

				//get the field maps
				XmlNodeList xmlNlFields = x.SelectNodes( "//MAP/TABLE[@NAME='BIG_SEARCH_TABLE']/FIELD" );
				//create the table
				DShowProgressEvent( "create " + tblName, false );
				CreateDicTable( dbCon, tblName, xmlNlFields, true );

				//create a column list for the insert
				cols = "(";
				foreach( XmlNode xmlNField in xmlNlFields )
				{
					cols += xmlNField.Attributes["NAME"].Value.ToString() + ", ";
				}
				cols = cols.Substring( 0, ( cols.Length - 2 ) ) + ")";

				//get db type
				IMEDDataAccess.ConnectionType cType = IMEDDataAccess.CalculateConnectionType( dbCon );

				//get sql string
				switch( cType )
				{
					case IMEDDataAccess.ConnectionType.SQLServer:
						sql = x.SelectSingleNode( "//MAP/SQL/SQL" ).InnerText.ToString();
						break;
					case IMEDDataAccess.ConnectionType.Oracle:
						sql = x.SelectSingleNode( "//MAP/SQL/ORA" ).InnerText.ToString();
						break;
				}
			
				//get big table data from all tables
				DShowProgressEvent( "get " + tblName + " data", false );
				ds = DataAccess.GetDataSet( dbCon, sql );
				bigCount = ds.Tables[0].Rows.Count;

				//add each row to big table
				foreach( DataRow row in ds.Tables[0].Rows )
				{
					string vals = "VALUES (";
					for( int n = 0; n < ds.Tables[0].Columns.Count; n++ )
					{
						vals += GetSqlValue(row[n].ToString(), 
							System.Convert.ToInt32( xmlNlFields[n].Attributes["TYPE"].Value.ToString() ) ) + ", ";
					}
					vals = vals.Substring( 0, ( vals.Length - 2 ) ) + ")";

					DShowProgressEvent( "insert BIG_SEARCH_TABLE, " + vals, false );
					sql = "INSERT INTO " + tblName + " " + cols + " " + vals;
					DataAccess.RunSQL( dbCon, sql );
				}
			}
			finally
			{
				ds.Dispose();
			}
		}

		/// <summary>
		/// Is search string in arraylist
		/// </summary>
		/// <param name="al"></param>
		/// <param name="search"></param>
		/// <returns></returns>
		private static bool IsInArray( ArrayList al, string search )
		{
			for( int n = 0; n < al.Count; n++ )
			{
				if( al[n].ToString() == search ) return( true );
			}
			return( false );
		}
	}
}
