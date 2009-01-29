using System;
using System.Collections;
using System.Data;
using log4net;

namespace InferMed.MACRO.ClinicalCoding.MACROCCBS30
{
	/// <summary>
	/// Collection of clinical coding dictionaries
	/// </summary>
	public class Dictionaries
	{
		//logging
		private static readonly ILog log = LogManager.GetLogger( typeof( Dictionaries ) );

		private ArrayList _dictionaries = null;

		public Dictionaries()
		{
			//use Init() as constructor because this is compiled for use in a vb6 application
			//and com does not support arguments in constructors
			//when this object is no longer being used by vb6, this constructor can be replaced 
			//and overloaded by methods below
		}

		/// <summary>
		/// Initialiser
		/// </summary>
		/// <param name="con"></param>
		public void Init( string con )
		{
			DataSet ds = null;
			
			try
			{
				//get details of all imported dictionaries
				string sql = "SELECT * FROM DICTIONARIES ORDER BY DICTIONARYNAME, DICTIONARYVERSION";
				ds = CCDataAccess.GetDataSet( con, sql );

				_dictionaries = new ArrayList();
				//add them to the dictionary list
				for( int n = 0; n < ds.Tables[0].Rows.Count; n++ )
				{
					Dictionary d = new Dictionary();
					d.Init( System.Convert.ToInt32( ds.Tables[0].Rows[n]["DictionaryId"].ToString() ), ds.Tables[0].Rows[n]["DictionaryName"].ToString(),
						ds.Tables[0].Rows[n]["DictionaryVersion"].ToString(), ds.Tables[0].Rows[n]["DictionaryConnection"].ToString() );
					_dictionaries.Add( d );
				}
			}
			finally
			{
				ds.Dispose();
			}
		}
		
		/// <summary>
		/// How many dictionaries are in the collection
		/// </summary>
		public int Count
		{
			get 
			{ 
				if( _dictionaries == null ) 
				{
					return( 0 );
				}
				else
				{
					return( _dictionaries.Count );
				}
			}
		}

		/// <summary>
		/// Get a dictionary from an id
		/// </summary>
		/// <param name="id"></param>
		/// <returns></returns>
		public Dictionary DictionaryFromId( int id )
		{
			Dictionary d = null;
			int n = 0;

			if( _dictionaries != null )
			{
				while( ( d == null ) && ( n < _dictionaries.Count ) )
				{
					Dictionary tempd = ( Dictionary )_dictionaries[n];
					if( tempd.Id == id ) d = tempd;
					n++;
				}
			}

			return( d );
		}

		/// <summary>
		/// Get a dictionary from a version
		/// </summary>
		/// <param name="dName"></param>
		/// <param name="dVersion"></param>
		/// <returns></returns>
		public Dictionary DictionaryFromVersion( string dName, string dVersion )
		{
			Dictionary d = null;
			int n = 0;

			if( _dictionaries != null )
			{
				while( ( d == null ) && ( n < _dictionaries.Count ) )
				{
					Dictionary tempd = ( Dictionary )_dictionaries[n];
					if( ( tempd.Name == dName ) && ( tempd.Version == dVersion ) ) d = tempd;
					n++;
				}
			}

			return( d );
		}

		/// <summary>
		/// Load a dictionary from the database
		/// </summary>
		/// <param name="con"></param>
		/// <param name="dName"></param>
		/// <param name="dVersion"></param>
		/// <returns></returns>
		public Dictionary DictionaryFromDatabase( string con, string dName, string dVersion )
		{
			DataSet ds = null;
			Dictionary d = null;
			
			try
			{
				string sql = "SELECT * FROM DICTIONARIES "
         + "WHERE DICTIONARYNAME = '" + dName + "' AND DICTIONARYVERSION = '" + dVersion + "' "
         + "ORDER BY DICTIONARYNAME, DICTIONARYVERSION";
				ds = CCDataAccess.GetDataSet( con, sql );

				if( ds.Tables[0].Rows.Count > 0 )
				{
					d = new Dictionary();
					d.Init( System.Convert.ToInt32( ds.Tables[0].Rows[0]["DictionaryId"].ToString() ), ds.Tables[0].Rows[0]["DictionaryName"].ToString(),
						ds.Tables[0].Rows[0]["DictionaryVersion"].ToString(), ds.Tables[0].Rows[0]["DictionaryConnection"].ToString() );
				}
				return( d );
			}
			finally
			{
				ds.Dispose();
			}
		}

		/// <summary>
		/// Get the list of dictionaries
		/// </summary>
		public ArrayList DictionaryList
		{
			get 
			{
				if( _dictionaries == null )
				{
					return( null );
				}
				else
				{
					return( _dictionaries );
				}
			}
		}

		/// <summary>
		/// What is the current maximum dictionary id
		/// </summary>
		private int MaxId
		{
			get
			{
				int id = 0;

				foreach( Dictionary tempd in _dictionaries )
				{
					if( tempd.Id > id ) id = tempd.Id;
				}

				return( id + 1 );
			}
		}

		/// <summary>
		/// Add a new dictionary
		/// </summary>
		/// <param name="con"></param>
		/// <param name="dName"></param>
		/// <param name="dVersion"></param>
		/// <param name="dConnection"></param>
		/// <param name="err"></param>
		/// <returns></returns>
		public void AddNew( string con, string dName, string dVersion, string dConnection )
		{
			Dictionary d = null;

			if( _dictionaries == null )
			{
				log.Info( "Initialising dictionaries" );
				Init( con );
			}

			if( DictionaryFromVersion( dName, dVersion ) != null )
			{
				log.Debug( "Dictionary exists " + dName + " " + dVersion );
				//dictionary with this name and version already exists
				Exception ex = new Exception( "Dictionary already exists" );
				throw ex;
			}
			else
			{
				string sql = "INSERT INTO DICTIONARIES (DICTIONARYID, DICTIONARYNAME, DICTIONARYVERSION, DICTIONARYCONNECTION) "
					+ "VALUES "
					+ "(" + MaxId + ", '" + dName + "', '" + dVersion + "', '" + dConnection + "')";
				d = new Dictionary();
				d.Init( MaxId, dName, dVersion, dConnection );
				_dictionaries.Add( d );
				log.Debug( sql );

				CCDataAccess.RunSQL( con, sql );
			}
		}
	}
}
