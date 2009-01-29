using System;
using System.Reflection;
using System.IO;

namespace InferMed.MACRO.ClinicalCoding.Interface
{
	/// <summary>
	/// Interface class between MACRO and plugins
	/// </summary>
	public class Plugin
	{
		private string _name;
		private string _version;
		private string _dllpath;
		private string _nameSpace;
		private string _custom;
		private Assembly _plugin;
		private Type _pluginType;

		public Plugin()
		{
			//use Init() as constructor because this is compiled for use in a vb6 application
			//and com does not support arguments in constructors
		}

		/// <summary>
		/// Initialisation method
		/// </summary>
		/// <param name="name"></param>
		/// <param name="version"></param>
		/// <param name="dllPath"></param>
		/// <param name="nameSpace"></param>
		/// <param name="custom"></param>
		public void Init( string name, string version, string dllPath, string nameSpace, string custom )
		{
			_name = name;
			_version = version;
			if( Path.IsPathRooted( dllPath ) )
			{
				_dllpath = dllPath;
			}
			else
			{
				_dllpath = ( new FileInfo( Path.GetDirectoryName( AppDomain.CurrentDomain.BaseDirectory ) ) + @"\" + dllPath );
			}
			_dllpath = dllPath;
			_nameSpace = nameSpace;
			_custom = custom;

			//load the plugin dll
			_plugin = Assembly.LoadFrom( _dllpath );
			//get the type of the dll
			_pluginType = _plugin.GetType( _nameSpace, true );
		}

		/// <summary>
		/// Code a response value
		/// </summary>
		/// <param name="responseValue"></param>
		/// <param name="codedValue"></param>
		public void Code( ref string responseValue, ref string codedValue )
		{
			MethodInfo codeMethod = _pluginType.GetMethod( "Code" );
			object methodInstance = Activator.CreateInstance( _pluginType );

			object[] parameterList = new object[5];
			parameterList[0] = _name;
			parameterList[1] = _version;
			parameterList[2] = _custom;
			parameterList[3] = responseValue;
			parameterList[4] = codedValue;

			codeMethod.Invoke( methodInstance, parameterList );

			responseValue = ( string )parameterList[3];
			codedValue = ( string )parameterList[4];
		}

		/// <summary>
		/// Get an array of terms matching a response value
		/// </summary>
		/// <param name="responseValue"></param>
		/// <param name="maxReturn"></param>
		/// <param name="totalMatches"></param>
		/// <returns></returns>
		public string[,] FindTerm( string responseValue, int maxReturn, ref int totalMatches )
		{
			MethodInfo codeMethod = _pluginType.GetMethod( "FindTerm" );
			object methodInstance = Activator.CreateInstance( _pluginType );

			object[] parameterList = new object[6];
			parameterList[0] = _name;
			parameterList[1] = _version;
			parameterList[2] = _custom;
			parameterList[3] = responseValue;
			parameterList[4] = maxReturn;
			parameterList[5] = totalMatches;

			object resultsObject = codeMethod.Invoke( methodInstance, parameterList );

			totalMatches = ( int )parameterList[5];

			return( ( string[,] ) resultsObject );
		}

		/// <summary>
		/// Get a text version of a coded value
		/// </summary>
		/// <param name="codedValue"></param>
		/// <returns></returns>
		public string ToText( string codedValue )
		{
			MethodInfo codeMethod = _pluginType.GetMethod( "ToText" );
			object methodInstance = Activator.CreateInstance( _pluginType );

			object[] parameterList = new object[4];
			parameterList[0] = _name;
			parameterList[1] = _version;
			parameterList[2] = _custom;
			parameterList[3] = codedValue;

			object codedObject = codeMethod.Invoke( methodInstance, parameterList );

			return( ( string )codedObject );
		}

		/// <summary>
		/// Get an HTML version of a coded value
		/// </summary>
		/// <param name="codedValue"></param>
		/// <returns></returns>
		public string ToHTML( string codedValue )
		{
			MethodInfo codeMethod = _pluginType.GetMethod( "ToHTML" );
			object methodInstance = Activator.CreateInstance( _pluginType );

			object[] parameterList = new object[4];
			parameterList[0] = _name;
			parameterList[1] = _version;
			parameterList[2] = _custom;
			parameterList[3] = codedValue;

			object codedObject = codeMethod.Invoke( methodInstance, parameterList );

			return( ( string )codedObject );
		}

		/// <summary>
		/// Get a single line text version of a coded value
		/// </summary>
		/// <param name="codedValue"></param>
		/// <returns></returns>
		public string ToSingleLineText( string codedValue )
		{
			MethodInfo codeMethod = _pluginType.GetMethod( "ToSingleLineText" );
			object methodInstance = Activator.CreateInstance( _pluginType );

			object[] parameterList = new object[4];
			parameterList[0] = _name;
			parameterList[1] = _version;
			parameterList[2] = _custom;
			parameterList[3] = codedValue;

			object codedObject = codeMethod.Invoke( methodInstance, parameterList );

			return( ( string )codedObject );
		}

		/// <summary>
		/// Get an XML tree version of a coded value
		/// </summary>
		/// <param name="codedValue"></param>
		/// <returns></returns>
		public string ToXmlTree( string codedValue )
		{
			MethodInfo codeMethod = _pluginType.GetMethod( "ToXmlTree" );
			object methodInstance = Activator.CreateInstance( _pluginType );

			object[] parameterList = new object[4];
			parameterList[0] = _name;
			parameterList[1] = _version;
			parameterList[2] = _custom;
			parameterList[3] = codedValue;

			object codedObject = codeMethod.Invoke( methodInstance, parameterList );

			return( ( string )codedObject );
		}

		/// <summary>
		/// Display a tree dialog for a coded value
		/// </summary>
		/// <param name="codedValue"></param>
		public void ToTree( string codedValue )
		{
			MethodInfo codeMethod = _pluginType.GetMethod( "ToTree" );
			object methodInstance = Activator.CreateInstance( _pluginType );

			object[] parameterList = new object[4];
			parameterList[0] = _name;
			parameterList[1] = _version;
			parameterList[2] = _custom;
			parameterList[3] = codedValue;

			codeMethod.Invoke( methodInstance, parameterList );
		}
	}
}
