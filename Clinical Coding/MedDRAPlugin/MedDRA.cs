using System;
using InferMed.MACRO.ClinicalCoding.Interface;
using System.Windows.Forms;
using log4net.Config;
using log4net;
using System.IO;

namespace InferMed.MACRO.ClinicalCoding.Plugins
{
	/// <summary>
	/// Clinical coding MedDRA plugin for MACRO
	/// </summary>
	public class MedDRA : InferMed.MACRO.ClinicalCoding.Interface.IPlugin
	{
		//logging
		private static readonly ILog log = LogManager.GetLogger( typeof( MedDRA ) );

		private CCDataServer _dServer = null;

		/// <summary>
		/// Constructor
		/// </summary>
		public MedDRA()
		{
			//initialise logging
			XmlConfigurator.Configure( new FileInfo( Path.GetDirectoryName( AppDomain.CurrentDomain.BaseDirectory ) 
				+ @"\log4netconfig.xml" ) );
		}

		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="custom"></param>
		public MedDRA( string custom )
		{
			//initialise logging
			XmlConfigurator.Configure( new FileInfo( Path.GetDirectoryName( AppDomain.CurrentDomain.BaseDirectory ) 
				+ @"\log4netconfig.xml" ) );

			string dictionary, ip, cCon, dCon, dPrefix, preferencesFile;
			CCXml.MedDRAUnwrapXmlCustom( custom, out dictionary, out ip, out cCon, out dCon, out dPrefix, out preferencesFile );
			_dServer = new CCDataServer( CCDataServer.DicType.MEDDRA, ip, cCon );
		}

		/// <summary>
		/// Set the MedDRA dictionary
		/// </summary>
		/// <param name="custom"></param>
		/// <returns></returns>
		public bool SetDictionary( string custom )
		{
			string dictionary, ip, cCon, dCon, dPrefix, preferencesFile;
			CCXml.MedDRAUnwrapXmlCustom( custom, out dictionary, out ip, out cCon, out dCon, out dPrefix, out preferencesFile );
			return( _dServer.SetDictionary( dictionary, dCon, dPrefix ) );
		}

		public CCDataServer.ServerStatus Status
		{
			get { return( _dServer.Status ); }
		}

		/// <summary>
		/// Destructor
		/// </summary>
		public void Dispose()
		{
			if( _dServer != null )
			{
				_dServer.Dispose();
			}
		}

		/// <summary>
		/// Autoencode a response value
		/// </summary>
		/// <param name="responseValue"></param>
		/// <param name="totalMatches"></param>
		/// <returns></returns>
		public string AutoEncode( string responseValue, ref int totalMatches )
		{
			string[,] results = null;
			string codedValue = "";

			//find a term match
			results = _dServer.FindTerm( responseValue, CCDataServer.SearchType.llt, 1, ref totalMatches );
			if( results != null )
			{
				//if the term is the best match
				if( MedDRATerm.IsBestMatch( results[0, (int)CCDataServer.ResultCol.autoencoder] ) )
				{
					codedValue = CCXml.MedDRAWrapXml( _dServer.Dictionary, responseValue, results[0, (int)CCDataServer.ResultCol.primary],
						results[0, (int)CCDataServer.ResultCol.current], "",
						results[0, (int)CCDataServer.ResultCol.socKey], results[0, (int)CCDataServer.ResultCol.soc],
						results[0, (int)CCDataServer.ResultCol.socAbbrev], results[0, (int)CCDataServer.ResultCol.hlgtKey],
						results[0, (int)CCDataServer.ResultCol.hlgt], results[0, (int)CCDataServer.ResultCol.hltKey],
						results[0, (int)CCDataServer.ResultCol.hlt], results[0, (int)CCDataServer.ResultCol.ptKey], 
						results[0, (int)CCDataServer.ResultCol.pt], results[0, (int)CCDataServer.ResultCol.lltKey], 
						results[0, (int)CCDataServer.ResultCol.llt] );
				}
			}
			return( codedValue );
		}

		/// <summary>
		/// Find a MedDRA term
		/// </summary>
		/// <param name="name"></param>
		/// <param name="version"></param>
		/// <param name="custom"></param>
		/// <param name="responseValue"></param>
		/// <param name="maxReturn"></param>
		/// <param name="totalMatches"></param>
		/// <returns></returns>
		public string[,] FindTerm( string name, string version, string custom, string responseValue, int maxReturn,	
			ref int totalMatches )
		{
			return( _dServer.FindTerm( responseValue, CCDataServer.SearchType.llt, maxReturn, ref totalMatches ) );
		}

		/// <summary>
		/// Get MedDRA historical matches
		/// </summary>
		/// <param name="custom"></param>
		/// <param name="term"></param>
		/// <returns></returns>
		public MedDRATerm[] TermMatches( string custom, string term )
		{
			return( MedDRATerm.Matches( custom, term ) );
		}

		/// <summary>
		/// Save a MedDRA historical match
		/// </summary>
		/// <param name="custom"></param>
		/// <param name="codedValue"></param>
		public void SaveTermMatch( string custom, string codedValue )
		{
			MedDRATerm.SaveMatch( custom, codedValue );
		}

		/// <summary>
		/// Put a coded path in an xml structure
		/// </summary>
		/// <param name="dictionary"></param>
		/// <param name="originalTerm"></param>
		/// <param name="primary"></param>
		/// <param name="current"></param>
		/// <param name="lastSearch"></param>
		/// <param name="socKey"></param>
		/// <param name="socName"></param>
		/// <param name="socAbbrev"></param>
		/// <param name="hlgtKey"></param>
		/// <param name="hlgtName"></param>
		/// <param name="hltKey"></param>
		/// <param name="hltName"></param>
		/// <param name="ptKey"></param>
		/// <param name="ptName"></param>
		/// <param name="lltKey"></param>
		/// <param name="lltName"></param>
		/// <returns></returns>
		public string WrapXml( string dictionary, string originalTerm, string primary, string current, string lastSearch,
			string socKey, string socName, string socAbbrev, string hlgtKey, string hlgtName, string hltKey, string hltName,
			string ptKey, string ptName, string lltKey, string lltName )
		{
			return( CCXml.MedDRAWrapXml( dictionary, originalTerm, primary, current, lastSearch, socKey, socName, socAbbrev,
				hlgtKey, hlgtName, hltKey, hltName, ptKey, ptName, lltKey, lltName ) );
		}

		/// <summary>
		/// Code a response value
		/// </summary>
		/// <param name="responseValue"></param>
		/// <param name="codedValue"></param>
		public void Code( string name, string version, string custom, ref string responseValue, ref string codedValue )
		{
			MedDRABrowser mb = new MedDRABrowser( custom, responseValue );

			if ( mb.Init() )
			{
				mb.ShowDialog();

				if ( mb._accepted )
				{
					codedValue = CCXml.MedDRAWrapXml( mb._dictionary, mb._originalTerm, mb._primary, mb._current, mb._lastSearch,
						mb._term._socKey, mb._term._soc, mb._term._socAbbrev, mb._term._hlgtKey, mb._term._hlgt, mb._term._hltKey,
						mb._term._hlt, mb._term._ptKey, mb._term._pt, mb._term._lltKey, mb._term._llt );
					SaveTermMatch( custom, codedValue  );
				}
			}
			mb.Dispose();
		}

		/// <summary>
		/// Function returns a text description from an XML coded path
		/// </summary>
		/// <param name="codedValue"></param>
		/// <returns></returns>
		public string ToText( string name, string version, string custom, string codedValue )
		{
			string dictionary, soc, hlgt, hlt, pt, llt;
			CCXml.MedDRAUnwrapXmlNames( codedValue, out dictionary, out soc, out hlgt, out hlt, out pt, out llt );
			return( "[SOC] : " + soc + " [HLGT] : " + hlgt + " [HLT] : " + hlt + " [PT] : " + pt + " [LLT] : " + llt );
		}

		/// <summary>
		/// Function returns an HTML description from an XML coded path
		/// </summary>
		/// <param name="codedValue"></param>
		/// <returns></returns>
		public string ToHTML( string name, string version, string custom, string codedValue )
		{
			string dictionary, soc, hlgt, hlt, pt, llt;
			CCXml.MedDRAUnwrapXmlNames( codedValue, out dictionary, out soc, out hlgt, out hlt, out pt, out llt );
			return( "[SOC] : " + soc + " [HLGT] : " + hlgt + " [HLT] : " + hlt + " [PT] : " + pt + " [LLT] : " + llt );
		}

		/// <summary>
		/// Function returns a single line description from an XML coded path
		/// </summary>
		/// <param name="codedValue"></param>
		/// <returns></returns>
		public string ToSingleLineText( string name, string version, string custom, string codedValue )
		{
			string dictionary, soc, hlgt, hlt, pt, llt;
			CCXml.MedDRAUnwrapXmlNames( codedValue, out dictionary, out soc, out hlgt, out hlt, out pt, out llt );
			return( "[SOC] : " + soc + " [HLGT] : " + hlgt + " [HLT] : " + hlt + " [PT] : " + pt + " [LLT] : " + llt );
		}

		/// <summary>
		/// Function returns the path tree from an XML coded path
		/// </summary>
		/// <param name="codedValue"></param>
		/// <returns></returns>
		public string ToXmlTree( string name, string version, string custom, string codedValue )
		{
			string tree;
			CCXml.MedDRAUnwrapXmlTree( codedValue, out tree );
			return( tree );
		}

		/// <summary>
		/// Function displays a tree dialog
		/// </summary>
		/// <param name="name"></param>
		/// <param name="version"></param>
		/// <param name="custom"></param>
		/// <param name="codedValue"></param>
		public void ToTree( string name, string version, string custom, string codedValue )
		{
			MedDRATree mt = new MedDRATree( codedValue );
			mt.ShowDialog();
			mt.Dispose();
		}
	}
}
