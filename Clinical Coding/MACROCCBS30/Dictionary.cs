using System;
using System.Xml;
using InferMed.MACRO.ClinicalCoding.Interface;

namespace InferMed.MACRO.ClinicalCoding.MACROCCBS30
{
	/// <summary>
	/// Clinical coding dictionary
	/// </summary>
	public class Dictionary
	{
		//xml connection string constants
		private const string _DICT_NODE = "//DICTIONARY";
		private const string _DICT_PLUGIN_NODE = "//DICTIONARY/PLUGIN";
		private const string _DICT_CUSTOM_NODE = "//DICTIONARY/CUSTOM";

		private const string _DICT_NAME_ATT = "NAME";
		private const string _DICT_VERSION_ATT = "VERSION";
		private const string _DICT_PLUGIN_NAMESPACE_ATT = "NAMESPACE";
		private const string _DICT_PLUGIN_PATH_ATT = "PATH";
		private const string _DICT_PLUGIN_DLLNAME_ATT = "DLLNAME";

		//dictionary id, name, version
		private int _dId;
		private string _dName;
		private string _dVersion;

		//dictionary xml connection string
		private string _dXmlCon;

		//subsets of dictionary xml connection string: namespace, directory path, file name, custom data
		private string _pluginNameSpace;
		private string _pluginPath;
		private string _pluginName;
		private string _pluginCustom;


		public Dictionary()
		{
			//use Init() as constructor because this is compiled for use in a vb6 application
			//and com does not support arguments in constructors
			//when this object is no longer being used by vb6, this constructor can be replaced 
			//and overloaded by methods below
		}

		/// <summary>
		/// Initialiser
		/// </summary>
		/// <param name="dId"></param>
		/// <param name="dName"></param>
		/// <param name="dVersion"></param>
		/// <param name="dXmlCon"></param>
		public void Init( int dId, string dName, string dVersion, string dXmlCon )
		{
			_dId = dId;
			_dName = dName;
			_dVersion = dVersion;
			_dXmlCon = dXmlCon;

			System.Xml.XmlDocument x = new XmlDocument();
			x.LoadXml( _dXmlCon );

			_pluginNameSpace = x.SelectSingleNode( _DICT_PLUGIN_NODE ).Attributes[ _DICT_PLUGIN_NAMESPACE_ATT ].Value.ToString();
			_pluginPath = x.SelectSingleNode( _DICT_PLUGIN_NODE ).Attributes[ _DICT_PLUGIN_PATH_ATT ].Value.ToString();
			_pluginName = x.SelectSingleNode( _DICT_PLUGIN_NODE ).Attributes[ _DICT_PLUGIN_DLLNAME_ATT ].Value.ToString();
			_pluginCustom = x.SelectSingleNode( _DICT_CUSTOM_NODE ).OuterXml.ToString();
		}

		/// <summary>
		/// Code a response using a dictionary
		/// </summary>
		/// <param name="responseValue"></param>
		/// <param name="codedValue"></param>
		public void Code( ref string responseValue, ref string codedValue )
		{
			Plugin p = new Plugin();
			p.Init( _dName, _dVersion, _pluginPath + _pluginName, _pluginNameSpace, _pluginCustom );
			p.Code( ref responseValue, ref codedValue );
		}

		/// <summary>
		/// Get a text version of a value coded by a dictionary
		/// </summary>
		/// <param name="codedValue"></param>
		/// <param name="text"></param>
		/// <param name="err"></param>
		/// <returns></returns>
		public bool ToText( string codedValue, ref string text, ref string err )
		{
			try
			{
				Plugin p = new Plugin();
				p.Init( _dName, _dVersion, _pluginPath + _pluginName, _pluginNameSpace, _pluginCustom );
				text = p.ToText( codedValue );
				return( true );
			}
			catch( Exception ex )
			{
				err = ex.Message + " " + ex.InnerException + "|Dictionary.ToText";
				return( false );
			}
		}

		/// <summary>
		/// Get an html version of a value coded by a dictionary
		/// </summary>
		/// <param name="codedValue"></param>
		/// <param name="html"></param>
		/// <param name="err"></param>
		/// <returns></returns>
		public bool ToHTML( string codedValue, ref string html, ref string err )
		{
			try
			{
				Plugin p = new Plugin();
				p.Init( _dName, _dVersion, _pluginPath + _pluginName, _pluginNameSpace, _pluginCustom );
				html = p.ToHTML( codedValue );
				return( true );
			}
			catch( Exception ex )
			{
				err = ex.Message + " " + ex.InnerException + "|Dictionary.ToHTML";
				return( false );
			}
		}

		/// <summary>
		/// Get an Xml tree version of a value coded by a dictionary
		/// </summary>
		/// <param name="codedValue"></param>
		/// <param name="xml"></param>
		/// <param name="err"></param>
		/// <returns></returns>
		public bool ToXmlTree( string codedValue, ref string xml, ref string err )
		{
			try
			{
				Plugin p = new Plugin();
				p.Init( _dName, _dVersion, _pluginPath + _pluginName, _pluginNameSpace, _pluginCustom );
				xml = p.ToText( codedValue );
				return( true );
			}
			catch( Exception ex )
			{
				err = ex.Message + " " + ex.InnerException + "|Dictionary.ToXmlTree";
				return( false );
			}
		}

		/// <summary>
		/// Display a dialog with a tree version of a value coded by a dictionary
		/// </summary>
		/// <param name="codedValue"></param>
		/// <param name="err"></param>
		/// <returns></returns>
		public bool ToTree( string codedValue, ref string err )
		{
			try
			{
				Plugin p = new Plugin();
				p.Init( _dName, _dVersion, _pluginPath + _pluginName, _pluginNameSpace, _pluginCustom );
				p.ToTree( codedValue );
				return( true );
			}
			catch( Exception ex )
			{
				err = ex.Message + " " + ex.InnerException + "|Dictionary.ToTree";
				return( false );
			}
		}

		public int Id
		{
			get { return( _dId ); }
		}

		public string Name
		{
			get { return( _dName ); }
		}

		public string Version
		{
			get { return( _dVersion ); }
		}

		public string Connection
		{
			get { return( _dXmlCon ); }
		}

		public string Custom
		{
			get { return( _pluginCustom ); }
		}
	}
}
