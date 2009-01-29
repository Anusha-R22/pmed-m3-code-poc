using System;
using System.Xml;
using System.IO;
using System.Text;

namespace InferMed.MACRO.ClinicalCoding.Plugins
{
	/// <summary>
	/// Clinical coding Xml functions
	/// </summary>
	public class CCXml
	{
		private CCXml()
		{
		}

		/// <summary>
		/// Get an Xml connection string
		/// </summary>
		/// <param name="name"></param>
		/// <param name="version"></param>
		/// <param name="pluginNamespace"></param>
		/// <param name="pluginPath"></param>
		/// <param name="pluginDllName"></param>
		/// <param name="custom"></param>
		/// <returns></returns>
		public static string WrapXmlConnection( string name, string version, string pluginNamespace, string pluginPath,
			string pluginDllName, string custom )
		{
			//<DICTIONARY NAME="MedDRA" VERSION="1.5">
			//  <PLUGIN NAMESPACE="InferMed.MACRO.ClinicalCoding.Plugins.MedDRA" PATH="C:\MACRO30\PluginInterface\Plugins\MedDRA\bin\Release" DLLNAME="MedDRAPlugin.dll"/>
			//  <CUSTOM> ... </CUSTOM>
			//</DICTIONARY>


			System.IO.MemoryStream s = new MemoryStream();
			System.Xml.XmlTextWriter tw = new XmlTextWriter( s, System.Text.Encoding.Unicode );
			System.Text.UnicodeEncoding e = new UnicodeEncoding();

			
			tw.WriteStartElement( "DICTIONARY" );
			tw.WriteAttributeString( "NAME", name );
			tw.WriteAttributeString( "VERSION", version );

			tw.WriteStartElement( "PLUGIN" );
			tw.WriteAttributeString( "NAMESPACE", pluginNamespace );
			tw.WriteAttributeString( "PATH", pluginPath );
			tw.WriteAttributeString( "DLLNAME", pluginDllName );
			tw.WriteEndElement();

			tw.WriteRaw( custom );
			
			tw.WriteEndElement();

			tw.Close();

			return e.GetString( s.ToArray()).Substring(1 );
		}

		/// <summary>
		/// Get an Xml custom string
		/// </summary>
		/// <param name="dictionary"></param>
		/// <param name="ip"></param>
		/// <param name="cCon"></param>
		/// <param name="dCon"></param>
		/// <param name="dPrefix"></param>
		/// <param name="preferencesPath"></param>
		/// <param name="preferencesFile"></param>
		/// <returns></returns>
		public static string MedDRAWrapXmlCustom( string dictionary, string ip, string cCon, string dCon, string dPrefix, 
			string preferencesPath, string preferencesFile )
		{
			//<CUSTOM>
			//  <DICTIONARY NAME="MedDRA V1.5" />
			//  <SERVER IP="" />
			//  <CDATABASE CONNECTION="PROVIDER=SQLOLEDB;DATA SOURCE=IAN;DATABASE=SQL_MEDICODER;USER ID=sa;PASSWORD=macrotm;" />
			//  <DDATABASE CONNECTION="PROVIDER=SQLOLEDB;DATA SOURCE=IAN;DATABASE=SQL_MEDICODER;USER ID=sa;PASSWORD=macrotm;" PREFIX="V15_" />
			//  <PREFERENCES PATH="E:\vss\MACRO30\PluginInterface\Plugins\MedDRA\bin\Release" FILE="MedDRAPluginPreferences.txt" />
			//</CUSTOM></DICTIONARY>

			System.IO.MemoryStream s = new MemoryStream();
			System.Xml.XmlTextWriter tw = new XmlTextWriter( s, System.Text.Encoding.Unicode );
			System.Text.UnicodeEncoding e = new UnicodeEncoding();


			tw.WriteStartElement( "CUSTOM" );

			tw.WriteStartElement( "DICTIONARY" );
			tw.WriteAttributeString( "NAME", dictionary );
			tw.WriteEndElement();
				
			tw.WriteStartElement( "SERVER" );
			tw.WriteAttributeString( "IP", ip );
			tw.WriteEndElement();
			
			tw.WriteStartElement( "CDATABASE" );
			tw.WriteAttributeString( "CONNECTION", cCon );
			tw.WriteEndElement();

			tw.WriteStartElement( "DDATABASE" );
			tw.WriteAttributeString( "CONNECTION", dCon );
			tw.WriteAttributeString( "PREFIX", dPrefix );
			tw.WriteEndElement();

			tw.WriteStartElement( "PREFERENCES" );
			tw.WriteAttributeString( "PATH", preferencesPath );
			tw.WriteAttributeString( "FILE", preferencesFile );
			tw.WriteEndElement();

			tw.WriteEndElement();

			tw.Close();

			return e.GetString( s.ToArray()).Substring(1 );
		}

		/// <summary>
		/// Unwrap custom xml string
		/// </summary>
		/// <param name="custom"></param>
		/// <param name="dictionary"></param>
		/// <param name="ip"></param>
		/// <param name="cCon"></param>
		/// <param name="dCon"></param>
		/// <param name="dPrefix"></param>
		/// <param name="preferencesFile"></param>
		public static void MedDRAUnwrapXmlCustom( string custom, out string dictionary, out string ip, out string cCon, 
			out string dCon, out string dPrefix, out string preferencesFile )
		{
			System.Xml.XmlDocument x = new XmlDocument();

			x.LoadXml( custom );

			dictionary = x.SelectSingleNode( "//CUSTOM/DICTIONARY" ).Attributes["NAME"].Value.ToString();
			ip = x.SelectSingleNode( "//CUSTOM/SERVER" ).Attributes["IP"].Value.ToString();
			cCon = x.SelectSingleNode( "//CUSTOM/CDATABASE" ).Attributes["CONNECTION"].Value.ToString();
			dCon = x.SelectSingleNode( "//CUSTOM/DDATABASE" ).Attributes["CONNECTION"].Value.ToString();
			dPrefix = x.SelectSingleNode( "//CUSTOM/DDATABASE" ).Attributes["PREFIX"].Value.ToString();
			preferencesFile = x.SelectSingleNode( "//CUSTOM/PREFERENCES" ).Attributes["PATH"].Value.ToString() + @"\" +
				x.SelectSingleNode( "//CUSTOM/PREFERENCES" ).Attributes["FILE"].Value.ToString();
		}

		/// <summary>
		/// Extract the path tree from an Xml coded path
		/// </summary>
		/// <param name="codedValue"></param>
		/// <param name="tree"></param>
		public static void MedDRAUnwrapXmlTree( string codedValue, out string tree )
		{
			System.Xml.XmlDocument x = new XmlDocument();

			x.LoadXml( codedValue );

			tree = x.SelectSingleNode( "//CODINGDETAILS/SOC" ).OuterXml.ToString();
		}

		/// <summary>
		/// Unwrap the original term from an Xml coded apth
		/// </summary>
		/// <param name="codedValue"></param>
		/// <param name="term"></param>
		public static void MedDRAUnwrapXmlTerm( string codedValue, out string term )
		{
			System.Xml.XmlDocument x = new XmlDocument();

			x.LoadXml( codedValue );

			term = x.SelectSingleNode( "//CODINGDETAILS/ORIGINALTERM" ).InnerText.ToString();
		}

		/// <summary>
		/// Extract the keys from an Xml coded path
		/// </summary>
		/// <param name="codedValue"></param>
		/// <param name="dictionary"></param>
		/// <param name="soc"></param>
		/// <param name="hlgt"></param>
		/// <param name="hlt"></param>
		/// <param name="pt"></param>
		/// <param name="llt"></param>
		public static void MedDRAUnwrapXmlKeys( string codedValue, out string dictionary, out string soc, out string hlgt, out string hlt, 
			out string pt, out string llt )
		{
			System.Xml.XmlDocument x = new XmlDocument();

			x.LoadXml( codedValue );

			dictionary = x.SelectSingleNode( "//CODINGDETAILS/DICTIONARY" ).InnerText.ToString();
			soc = x.SelectSingleNode("//CODINGDETAILS/SOC" ).Attributes["KEY"].Value.ToString();
			hlgt = x.SelectSingleNode("//CODINGDETAILS/SOC/HLGT" ).Attributes["KEY"].Value.ToString();
			hlt = x.SelectSingleNode("//CODINGDETAILS/SOC/HLGT/HLT" ).Attributes["KEY"].Value.ToString();
			pt = x.SelectSingleNode("//CODINGDETAILS/SOC/HLGT/HLT/PT" ).Attributes["KEY"].Value.ToString();
			llt = x.SelectSingleNode("//CODINGDETAILS/SOC/HLGT/HLT/PT/LLT" ).Attributes["KEY"].Value.ToString();
		}

		/// <summary>
		/// Extract the names from an Xml coded path
		/// </summary>
		/// <param name="codedValue"></param>
		/// <param name="dictionary"></param>
		/// <param name="soc"></param>
		/// <param name="hlgt"></param>
		/// <param name="hlt"></param>
		/// <param name="pt"></param>
		/// <param name="llt"></param>
		public static void MedDRAUnwrapXmlNames( string codedValue, out string dictionary, out string soc, out string hlgt, out string hlt, 
			out string pt, out string llt )
		{
			System.Xml.XmlDocument x = new XmlDocument();

			x.LoadXml( codedValue );

			dictionary = x.SelectSingleNode( "//CODINGDETAILS/DICTIONARY" ).InnerText.ToString();
			soc = x.SelectSingleNode( "//CODINGDETAILS/SOC" ).Attributes["NAME"].Value.ToString();
			hlgt = x.SelectSingleNode( "//CODINGDETAILS/SOC/HLGT" ).Attributes["NAME"].Value.ToString();
			hlt = x.SelectSingleNode( "//CODINGDETAILS/SOC/HLGT/HLT" ).Attributes["NAME"].Value.ToString();
			pt = x.SelectSingleNode( "//CODINGDETAILS/SOC/HLGT/HLT/PT" ).Attributes["NAME"].Value.ToString();
			llt = x.SelectSingleNode( "//CODINGDETAILS/SOC/HLGT/HLT/PT/LLT" ).Attributes["NAME"].Value.ToString();
		}

		/// <summary>
		/// Create an Xml structure containing a coded path
		/// </summary>
		/// <param name="b"></param>
		/// <returns></returns>
		public static string MedDRAWrapXml( string dictionary, string originalTerm, string primary, string current, string lastSearch,
			string socKey, string socName, string socAbbrev, string hlgtKey, string hlgtName, string hltKey, string hltName,
			string ptKey, string ptName, string lltKey, string lltName )
		{
			//example of xml structure created
			//
			//			<CODINGDETAILS>
			//				<DICTIONARY>MedDRA V8.0</DICTIONARY> 
			//				<ORIGINALTERM>Headache</ORIGINALTERM> 
			//				<PRIMARY>Y</PRIMARY> 
			//				<CURRENT>Y</CURRENT> 
			//				<LASTSEARCH>Headache</LASTSEARCH> 
			//			- <SOC KEY="10029205" NAME="Nervous system disorders" ABBREV="Nerv">
			//				- <HLGT KEY="10019231" NAME="Headaches">
			//					- <HLT KEY="10019231" NAME="Headaches">
			//						- <PT KEY="10036313" NAME="Post-traumatic headache">
			//								<LLT KEY="10019222" NAME="Headache post-traumatic" /> 
			//							</PT>
			//						</HLT>
			//					</HLGT>
			//				</SOC>
			//			</CODINGDETAILS>

			System.IO.MemoryStream s = new MemoryStream();
			System.Xml.XmlTextWriter tw = new XmlTextWriter( s, System.Text.Encoding.Unicode );
			System.Text.UnicodeEncoding e = new UnicodeEncoding();
			
			
			tw.WriteStartDocument();
			tw.WriteStartElement( "CODINGDETAILS" );

			tw.WriteElementString( "DICTIONARY", dictionary );
			tw.WriteElementString( "ORIGINALTERM", originalTerm.ToUpper() );
//			tw.WriteElementString("PRIMARY", primary);
//			tw.WriteElementString("CURRENT", current);
//			tw.WriteElementString("LASTSEARCH", lastSearch);

			tw.WriteStartElement( "SOC" );
			tw.WriteAttributeString( "KEY", socKey );
			tw.WriteAttributeString( "NAME", MedDRATerm.Clean( socName ) );
			tw.WriteAttributeString( "ABBREV", MedDRATerm.Clean( socAbbrev ) );
			
			tw.WriteStartElement( "HLGT" );
			tw.WriteAttributeString( "KEY", hlgtKey );
			tw.WriteAttributeString( "NAME", MedDRATerm.Clean( hlgtName ) );
			
			tw.WriteStartElement( "HLT" );
			tw.WriteAttributeString( "KEY", hltKey );
			tw.WriteAttributeString( "NAME", MedDRATerm.Clean( hltName ) );

			tw.WriteStartElement( "PT" );
			tw.WriteAttributeString( "KEY", ptKey );
			tw.WriteAttributeString( "NAME", MedDRATerm.Clean( ptName ) );
			
			tw.WriteStartElement( "LLT" );
			tw.WriteAttributeString( "KEY", lltKey );
			tw.WriteAttributeString( "NAME", MedDRATerm.Clean( lltName ) );
			tw.WriteEndElement();

			tw.WriteEndElement();
			tw.WriteEndElement();
			tw.WriteEndElement();
			tw.WriteEndElement();
			tw.WriteEndElement();

			tw.WriteEndDocument();
			tw.Close();

			return e.GetString( s.ToArray()).Substring(1 );
		}
	}
}
