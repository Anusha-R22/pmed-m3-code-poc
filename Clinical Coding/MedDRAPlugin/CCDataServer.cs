using System;
using System.Threading;
using log4net;

namespace InferMed.MACRO.ClinicalCoding.Plugins
{
	/// <summary>
	/// Encapsulates the dictionary searching object
	/// </summary>
	public class CCDataServer
	{
		//logging
		private static readonly ILog log = LogManager.GetLogger( typeof( CCDataServer ) );

		//socketserver object
		private SocketXExeM.clsSocket _socket = null;
		
		//socketserver initialisation parameters
		private object _licProductName = "Browser6560DP5924";
		private object _licProductId = "";
		private object _licSection1 = "";
		private object _licSection2 = "";
		private object _licSystemId = "";
		private object _licPassword1 = "";
		private object _licPassword2 = "";
		private object _intFirst = 0;
		private object _hWind = 0;
		private int _lAPICode = 4322;
		//private short _lmsID = 1;

		//socketserver tables and sql
		private string _TDERIVATIVES = "BE_DERIVATIVES";
		private string _TINVALID = "BE_INVALID_WORDS";
		private string _TWEIGHTS = "SELECT * FROM BE_WEIGHTING ORDER BY STEP_NO ASC";

		//socketserver language
		private string _lang = "ENGLISH";

		//maximum records to fetch in a single operation
		private int _MAXFETCH = 500;

		//socketserver search parameters
		private int _slotType = 0;
		private int _slots = 100;
		private int _session = 0;

		//number of result columns returned from socketserver
		private int _RESULTCOLS = 24;

		//socketserver result columns
		public enum ResultCol
		{
			lltKey = 0, llt = 1,
			ptKey = 2, pt = 3,
			hltKey = 4, hlt = 5,
			hlgtKey = 6, hlgt = 7,
			socKey = 8, soc = 9, socAbbrev = 10,
			primary = 17, current = 18,
			autoencoder = 19,
			fullMatch = 20, partMatch = 21, weight = 22
		}

		//socketserver search types
		public enum SearchType
		{
			llt = 4354, pt = 8450, hlt = 12546, hlgt = 16642, soc = 20738
		}

		//dictionary types
		public enum DicType
		{
			MEDDRA = 0, COSTART = 1, ICD9 = 2, ICD9CM = 3,  WHODRUG = 4, WHOART = 5
		}

		//connection status
		private enum ConStatus
		{
			Connected = 7
		}

		//state of readiness
		public enum ServerStatus
		{
			Disconnected, Connected, Ready
		}

		//connection parameters
		private const int _SLEEPTIME = 1000;
		private const int _MAXTRIES = 10;
		
		//private properties
		private string _dictionary = "";
		private string _ip = "";
		private string _cCon = "";
		private string _dCon = "";
		private string _dPrefix = "";
		private DicType _dType;
		private string _error = "";
		private ServerStatus _status = ServerStatus.Disconnected;


		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="dt"></param>
		/// <param name="dictionary"></param>
		/// <param name="ip"></param>
		/// <param name="con"></param>
		public CCDataServer( DicType dt, string dictionary, string ip, string cCon, string dCon, string dPrefix )
		{
			_dType = dt;
			_dictionary = dictionary;
			_ip = ip;
			_cCon = cCon;
			_dCon = dCon;
			_dPrefix = dPrefix;

			log.Debug( "dt=" + dt.ToString() + " dictionary=" + dictionary + " ip=" + ip + " dPrefix=" + dPrefix );

			//create new socketserver instance
			_socket = new SocketXExeM.clsSocket();

			//connect and initialise
			if( ConnectToServer() )
			{
				InitialiseDictionary();
			}
		}

		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="dt"></param>
		/// <param name="ip"></param>
		/// <param name="cCon"></param>
		public CCDataServer( DicType dt, string ip, string cCon )
		{
			_dType = dt;
			_ip = ip;
			_cCon = cCon;

			log.Debug( "dt=" + dt.ToString() + " ip=" + ip );

			//create new socketserver instance
			_socket = new SocketXExeM.clsSocket();

			//connect
			ConnectToServer();
		}

		/// <summary>
		/// Destructor
		/// </summary>
		public void Dispose()
		{
			try
			{
				log.Info( "Disposing" );
				//_socket.CloseConnection( ref _lAPICode, ref _lmsID );
				_socket.Quit( ref _lAPICode );
			}
			catch
			{
				//ignore any errors when quitting
			}
		}

		public string Error
		{
			get { return( _error ); }
		}

		public ServerStatus Status
		{
			get { return( _status ); }
		}

		/// <summary>
		/// Set the dictionary to use
		/// </summary>
		/// <param name="dictionary"></param>
		/// <param name="dCon"></param>
		/// <param name="dPrefix"></param>
		/// <returns></returns>
		public bool SetDictionary( string dictionary, string dCon, string dPrefix )
		{
			_dictionary = dictionary;
			_dCon = dCon;
			_dPrefix = dPrefix;

			log.Debug( "dictionary=" + dictionary + " PREFIX=" + dPrefix );
			if( InitialiseDictionary() )
			{
				log.Info( "Initialised" );
				return( LoadDictionary() );
			}
			else
			{
				return( false );
			}
		}

		public string Dictionary
		{
			get { return( _dictionary ); }
		}

		/// <summary>
		/// Connects to dictionary search engine via socketserver
		/// </summary>
		/// <returns></returns>
		private bool ConnectToServer()
		{
			int tries = 0;
			object s1 = "", s2 = "", l1 = 0, l2 = 0;
			
			try
			{
				_status = ServerStatus.Disconnected;

				if ( _ip != "" )
				{
					//if connecting to a remote server
					_socket.IPAddress = _ip;
					_socket.bServer = false;
					_socket.bStandalone = false;

					//try _MAXTRIES to connect to remote server
					_socket.Connect( ref _lAPICode );
					while ( ( !SocketConnected( _socket ) ) && ( tries <= _MAXTRIES ) )
					{
						Thread.Sleep( _SLEEPTIME );
						tries++;
					}
				}
				else
				{
					//if connecting to a local server
					_socket.bServer = false;
					_socket.bStandalone = true;
				}

				if ( ( _socket.bStandalone ) || ( SocketConnected( _socket ) ) )
				{
					log.Info( "Activating licence" );
					//if successfully connected
					_socket.ActivateLicense( ref _licProductName, ref _licProductId, ref _licSection1, ref _licSection2, ref _licSystemId,
						ref _licPassword1, ref _licPassword2, ref _intFirst, ref _hWind, ref s1, ref s2, ref l1, ref l2 );
					_status = ServerStatus.Connected;
					return( true );
				}
				else
				{
					log.Info( "Connection failed" );
					_error = "Unable to connect to server";
					return( false );
				}
			}
			catch( Exception ex )
			{
				log.Error( ex.Message + " : " + ex.InnerException );
				//if exception, store the message for retrieval
				_error = ex.Message + " : " + ex.InnerException;
				return( false );
			}
		}

		private bool InitialiseDictionary()
		{
			int res = 0, dType = ( int )_dType;

			try
			{
				if( ( _status == ServerStatus.Connected ) || ( _status == ServerStatus.Ready) )
				{
					_status = ServerStatus.Connected;

					if( ( _dictionary != "" ) && ( _dCon != "" ) )
					{
						log.Info( "Loading dictionary" );
						res = _socket.LoadDerivatives( ref _lAPICode, ref _cCon, ref dType, ref _TDERIVATIVES, ref _lang, ref _dictionary );
						res = _socket.LoadInvalidWords( ref _lAPICode, ref _cCon, ref dType, ref _TINVALID, ref _lang, ref _dictionary );
						res = _socket.LoadWeights( ref _lAPICode, ref _cCon, ref dType, ref _TWEIGHTS, ref _lang, ref _dictionary );
						
						//if successfully initialised
						_status = ServerStatus.Ready;
						return( true );
					}
					else
					{
						log.Info( "No dictionary" );
						_error = "Server cannot be initialised while no dictionary is set";
						return( false );
					}
				}
				else
				{
					log.Info( "Server disconnected" );
					_error = "Server cannot be initialised while disconnected";
					return( false );
				}
			}
			catch( Exception ex )
			{
				log.Error( ex.Message + " : " + ex.InnerException );
				//if exception, store the message for retrieval
				_error = ex.Message + " : " + ex.InnerException;
				return( false );
			}
		}

		/// <summary>
		/// Is the socketserver connected to the server
		/// </summary>
		/// <param name="socket"></param>
		/// <returns></returns>
		private bool SocketConnected( SocketXExeM.clsSocket socket )
		{
			short s = 0;
			return( ( ( int )socket.get_Status( ref s ) == ( int )ConStatus.Connected ) );
		}

		/// <summary>
		/// Open a server session
		/// </summary>
		private void OpenSession()
		{
			_session = _socket.AllocateSession( ref _lAPICode );
		}

		/// <summary>
		/// Close a server session
		/// </summary>
		private void CloseSession()
		{
			_socket.CloseSession( ref _lAPICode, ref _session );
			_session = 0;
		}

		/// <summary>
		/// Loads the plugin dictionary into memory
		/// </summary>
		/// <returns></returns>
		public bool LoadDictionary()
		{
			int dType = ( int )_dType, res = 0;
			bool loadOK = false;

			//note - this does not check for a change in database
			if( _socket.IsDictionaryLoaded( ref _lAPICode, ref dType, ref _lang, ref _dictionary ) != 1 )
			{
				log.Info( "Loading dictionary" );
				//load dictionary
				res = _socket.LoadDataCache( ref _lAPICode, ref _dCon, ref dType, ref _lang, ref _dictionary, ref _dPrefix );
				if( res > 0 )
				{
					loadOK = true;
				}
			}
			else
			{
				log.Info( "Dictionary already loaded" );
				loadOK = true;
			}

			return ( loadOK );
		}

		/// <summary>
		/// Prepare a search term for passing to the server
		/// </summary>
		/// <param name="term"></param>
		private void CleanTerm( ref string term )
		{
			term = term.Trim();
			term = term.ToUpper();
		}

		/// <summary>
		/// Search for a term and return the results
		/// </summary>
		/// <param name="term"></param>
		/// <returns></returns>
		public string[,] FindTerm( string term, SearchType searchType, int maxReturn, ref int totalMatches )
		{
			int dType = ( int )_dType, sType = ( int )searchType, matches = 0, matchesFetched = 0, 
				matchesToFetch = 0;
			string[,] results2D = null;
			
			CleanTerm( ref term );
			log.Debug( "term=" + term + " SearchType=" + searchType.ToString() + " maxReturn=" + maxReturn );

			if( term != "" )
			{
				if( LoadDictionary() )
				{
					//open server session
					OpenSession();

					log.Info( "Finding term" );
					//search for term matches
					totalMatches = _socket.FindTerm( ref _lAPICode, ref term, ref dType, ref _lang, ref _dictionary, ref sType, 
						ref _slotType, ref _slots, ref _session );
					log.Debug( "totalMatches=" + totalMatches );

					if( totalMatches > 0 )
					{
						//only get the maximum number of results
						matches = ( ( totalMatches > maxReturn ) && ( maxReturn > 0 ) ) ? maxReturn : totalMatches;

						//initialise a 2 dimensional array for the data
						results2D = new string[matches, _RESULTCOLS];

						log.Info( "Getting matches" );
						while( matchesFetched < matches )
						{
							matchesToFetch = ( matches - matchesFetched > _MAXFETCH ) ? _MAXFETCH : matches - matchesFetched;

							//create a string result array big enough to hold the results
							string[] results1D = new string[matchesToFetch * _RESULTCOLS + 1];
							//convert to an object
							object results1Do = ( object )results1D;

							//fetch the matches
							int recs = _socket.FetchMatch( ref _lAPICode, ref dType, ref _lang, ref _dictionary, ref matchesToFetch, 
								ref _session, ref results1Do );

							//if the matches were returned
							if( recs > 0 )
							{
								//convert back to a string array
								results1D = ( string[] )results1Do;
								//add to results
								AddToResults(ref results2D, results1D, matchesFetched, recs );
							}

							matchesFetched+=matchesToFetch;
						}
					}

					//close the server session
					CloseSession();
				}
			}
			return( results2D );
		}

		/// <summary>
		/// Add 1 dimensional results to 2 dimensional array
		/// </summary>
		/// <param name="results2D"></param>
		/// <param name="results1D"></param>
		/// <param name="startRow"></param>
		/// <param name="recs"></param>
		private void AddToResults(ref string[,] results2D, string[] results1D, int startRow, int recs)
		{
			//fill the 2 dimensional array with the one dimensional data
			for(int n = 0; n < recs; n++)
			{
				results2D[startRow + n, ( int )ResultCol.lltKey] = results1D[n * _RESULTCOLS + ( int )ResultCol.lltKey + 1];
				results2D[startRow + n, ( int )ResultCol.llt] = results1D[n * _RESULTCOLS + ( int )ResultCol.llt + 1];
				results2D[startRow + n, ( int )ResultCol.ptKey] = results1D[n * _RESULTCOLS + ( int )ResultCol.ptKey + 1];
				results2D[startRow + n, ( int )ResultCol.pt] = results1D[n * _RESULTCOLS + ( int )ResultCol.pt + 1];
				results2D[startRow + n, ( int )ResultCol.hltKey] = results1D[n * _RESULTCOLS + ( int )ResultCol.hltKey + 1];
				results2D[startRow + n, ( int )ResultCol.hlt] = results1D[n * _RESULTCOLS + ( int )ResultCol.hlt + 1];
				results2D[startRow + n, ( int )ResultCol.hlgtKey] = results1D[n * _RESULTCOLS + ( int )ResultCol.hlgtKey + 1];
				results2D[startRow + n, ( int )ResultCol.hlgt] = results1D[n * _RESULTCOLS + ( int )ResultCol.hlgt + 1];
				results2D[startRow + n, ( int )ResultCol.socKey] = results1D[n * _RESULTCOLS + ( int )ResultCol.socKey + 1];
				results2D[startRow + n, ( int )ResultCol.soc] = results1D[n * _RESULTCOLS + ( int )ResultCol.soc + 1];
				results2D[startRow + n, ( int )ResultCol.socAbbrev] = results1D[n * _RESULTCOLS + ( int )ResultCol.socAbbrev + 1];
				results2D[startRow + n, ( int )ResultCol.primary] = results1D[n * _RESULTCOLS + ( int )ResultCol.primary + 1];
				results2D[startRow + n, ( int )ResultCol.current] = results1D[n * _RESULTCOLS + ( int )ResultCol.current + 1];
				results2D[startRow + n, ( int )ResultCol.autoencoder] = results1D[n * _RESULTCOLS + ( int )ResultCol.autoencoder + 1];
				results2D[startRow + n, ( int )ResultCol.fullMatch] = results1D[n * _RESULTCOLS + ( int )ResultCol.fullMatch + 1];
				results2D[startRow + n, ( int )ResultCol.partMatch] = results1D[n * _RESULTCOLS + ( int )ResultCol.partMatch + 1];
				results2D[startRow + n, ( int )ResultCol.weight] = results1D[n * _RESULTCOLS + ( int )ResultCol.weight + 1];
			}
		}
	}
}
