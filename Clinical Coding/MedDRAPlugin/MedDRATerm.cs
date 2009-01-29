using System;
using System.Data;
using InferMed.Components;
using log4net;

namespace InferMed.MACRO.ClinicalCoding.Plugins
{
	/// <summary>
	/// MedDRA term object
	/// </summary>
	public class MedDRATerm
	{
		//logging
		private static readonly ILog log = LogManager.GetLogger( typeof( MedDRATerm ) );

		//forbidden characters
		public const string _FORBIDDEN_CHARS = "`¬|~\"";
		//path properties
		public string _socKey = "";
		public string _soc = "";
		public string _socAbbrev = "";
		public string _hlgtKey = "";
		public string _hlgt = "";
		public string _hltKey = "";
		public string _hlt = "";
		public string _ptKey = "";
		public string _pt = "";
		public string _lltKey = "";
		public string _llt = "";

		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="socKey"></param>
		/// <param name="soc"></param>
		/// <param name="socAbbrev"></param>
		/// <param name="hlgtKey"></param>
		/// <param name="hlgt"></param>
		/// <param name="hltKey"></param>
		/// <param name="hlt"></param>
		/// <param name="ptKey"></param>
		/// <param name="pt"></param>
		/// <param name="lltKey"></param>
		/// <param name="llt"></param>
		public MedDRATerm( string socKey, string soc, string socAbbrev, string hlgtKey, string hlgt, string hltKey, 
			string hlt, string ptKey, string pt, string lltKey, string llt )
		{
			_socKey = socKey;
			_soc = soc;
			_socAbbrev = socAbbrev;
			_hlgtKey = hlgtKey;
			_hlgt = hlgt;
			_hltKey = hltKey;
			_hlt = hlt;
			_ptKey = ptKey;
			_pt = pt;
			_lltKey = lltKey;
			_llt = llt;
		}

		/// <summary>
		/// Prepare a string for adding to a SQL string
		/// </summary>
		/// <param name="term"></param>
		/// <returns></returns>
		private static string SQLTerm( string term )
		{
			return( term.Replace( "'", "''" ).ToUpper() );
		}

		/// <summary>
		/// Get historical matches
		/// </summary>
		/// <param name="custom"></param>
		/// <param name="term"></param>
		/// <returns></returns>
		public static MedDRATerm[] Matches( string custom, string term )
		{
			string dictionary, ip, cCon, dCon, dPrefix, preferencesFile;
			CCXml.MedDRAUnwrapXmlCustom( custom, out dictionary, out ip, out cCon, out dCon, out dPrefix, out preferencesFile );
			return( Matches( dCon, dictionary, term ) );
		}

		/// <summary>
		/// Get historical matches
		/// </summary>
		/// <param name="con"></param>
		/// <param name="dictionary"></param>
		/// <param name="term"></param>
		/// <returns></returns>
		public static MedDRATerm[] Matches( string con, string dictionary, string term )
		{
			DataSet ds = null;
			
			try
			{
				log.Info( "FINDING HISTORICAL MATCHES" );
				string sql = "SELECT SOC_CODE, HLGT_CODE, HLT_CODE, PT_CODE, LLT_CODE "
					+ "FROM BE_HISTORY "
					+ "WHERE MEDDRA_VERSION = '" + dictionary + "' AND ORIGINAL_TERM = '" + SQLTerm( term ) + "'";
				log.Debug( "SQL=" + sql );
				ds = GetDataSet( con, sql );
				MedDRATerm[] historicalMatches = new MedDRATerm[ds.Tables[0].Rows.Count];

				log.Info( "PROCESSING HISTORICAL MATCHES" );
				for( int n = 0; n < ds.Tables[0].Rows.Count; n++ )
				{
					historicalMatches[n] = new MedDRATerm( ds.Tables[0].Rows[n]["SOC_CODE"].ToString(), "", "",
						ds.Tables[0].Rows[n]["HLGT_CODE"].ToString(), "", ds.Tables[0].Rows[n]["HLT_CODE"].ToString(), "",
						ds.Tables[0].Rows[n]["PT_CODE"].ToString(), "", ds.Tables[0].Rows[n]["LLT_CODE"].ToString(), "");
				}

				return( historicalMatches );
			}
			finally
			{
				ds.Dispose();
			}
		}

		/// <summary>
		/// Save a historical match
		/// </summary>
		/// <param name="custom"></param>
		/// <param name="codedValue"></param>
		public static void SaveMatch( string custom, string codedValue )
		{
			DataSet ds = null;

			try
			{
				string term, dictionary, socKey, hlgtKey, hltKey, ptKey, lltKey, ip, cCon, dCon, dPrefix, preferencesFile;
				CCXml.MedDRAUnwrapXmlCustom( custom, out dictionary, out ip, out cCon, out dCon, out dPrefix, out preferencesFile );
				CCXml.MedDRAUnwrapXmlKeys( codedValue, out dictionary, out socKey, out hlgtKey, out hltKey, out ptKey, out lltKey );
				CCXml.MedDRAUnwrapXmlTerm( codedValue, out term );

				string sql = "SELECT COUNT(*) AS COUNT "
					+ "FROM BE_HISTORY "
					+ "WHERE MEDDRA_VERSION = '" + dictionary + "' AND ORIGINAL_TERM = '" + term + "' "
					+ "AND SOC_CODE = '" + socKey + "' AND HLGT_CODE = '" + hlgtKey + "' AND HLT_CODE = '" + hltKey + "' "
					+ "AND PT_CODE = '" + ptKey + "' AND LLT_CODE = '" + lltKey + "'";
				log.Debug( "SQL=" + sql );
				ds = GetDataSet( dCon, sql );

				if( ds.Tables[0].Rows[0]["COUNT"].ToString() == "0" )
				{
					log.Info( "SAVING HISTORICAL MATCH" );
					sql = "INSERT INTO BE_HISTORY (SOC_CODE, HLGT_CODE, HLT_CODE, PT_CODE, LLT_CODE, MEDDRA_VERSION, ORIGINAL_TERM) "
						+ "VALUES ('" + socKey + "', '" + hlgtKey + "', '" + hltKey + "', '" + ptKey + "', '" + lltKey + "', "
						+ "'" + dictionary + "', '" + term + "')";
					log.Debug( "SQL=" + sql );
					RunSQL( dCon, sql );
				}
			}
			finally
			{
			}
		}

		/// <summary>
		/// Is this the best match
		/// </summary>
		/// <param name="autoencoder"></param>
		/// <returns></returns>
		public static bool IsBestMatch( string autoencoder )
		{
			if( autoencoder.Length != 2 ) return false;
			if( autoencoder.Substring( 0, 1 ) != "Y" ) return false;
			return true;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="connectionString"></param>
		/// <param name="sql"></param>
		/// <returns></returns>
		private static DataSet GetDataSet(string connectionString, string sql)
		{
			IMEDDataAccess imedData = new IMEDDataAccess();
			IDbConnection dbConn = null;
			
			try
			{
				// open ImedDataAccess class
				IMEDDataAccess.ConnectionType imedConnType;
				// calculate connection type
				imedConnType=IMEDDataAccess.CalculateConnectionType( connectionString );
				// get connection & open
				dbConn = imedData.GetConnection( imedConnType, connectionString );
				dbConn.Open();
				
				return( GetDataSet( dbConn, sql ) );
			}
			finally
			{
				//clean up objects
				dbConn.Close();
				dbConn.Dispose();
			}
		}

		/// <summary>
		/// Returns a dataset from the database
		/// </summary>
		/// <param name="dbConn">Database connection object</param>
		/// <param name="sql">SQL to run against the MACRO database</param>
		/// <returns>Dataset containing MACRO data</returns>
		private static DataSet GetDataSet(IDbConnection dbConn, string sql)
		{
			IMEDDataAccess imedData = new IMEDDataAccess();
			IDbCommand dbCommand = null;
			IDbDataAdapter dbDataAdapter = null;
			DataSet dataDB = new DataSet();

			try
			{
				// get command
				dbCommand = imedData.GetCommand( dbConn );
				// create command text
				dbCommand.CommandText = sql;
				// get data adaptor
				dbDataAdapter = imedData.GetDataAdapter( dbCommand );
				// fill dataset
				dbDataAdapter.Fill( dataDB );
				
				return( dataDB );
			}
			finally
			{
				//clean up objects
				dbCommand.Dispose();
			}
		}

		private static void RunSQL( string connectionString, string sql )
		{
			IMEDDataAccess imedData = new IMEDDataAccess();
			IDbConnection dbConn = null;
			IDbCommand dbCommand = null;

			try
			{
				// open ImedDataAccess class
				IMEDDataAccess.ConnectionType imedConnType;
				// calculate connection type
				imedConnType=IMEDDataAccess.CalculateConnectionType( connectionString );
				// get connection & open
				dbConn = imedData.GetConnection( imedConnType, connectionString );
				dbConn.Open();

				// get command
				dbCommand = imedData.GetCommand( dbConn );
				// create command text
				dbCommand.CommandText = sql;
				dbCommand.ExecuteNonQuery();
			}
			finally
			{
				//clean up objects
				dbCommand.Dispose();
			}
		}

		public static string Clean( string term )
		{
			string t = "";
			
			for( int n = 0; n < term.Length; n++ )
			{
				t += ( CharIsOK( term[n].ToString() ) ) ? term[n].ToString() : "";
			}
			return( t );
		}

		private static bool CharIsOK( string s )
		{
			for( int n = 0; n < _FORBIDDEN_CHARS.Length; n++ )
			{
				if( s == _FORBIDDEN_CHARS[n].ToString() ) return false;
			}
			return( true );
		}
	}
}
