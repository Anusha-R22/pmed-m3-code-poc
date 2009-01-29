using System;
using InferMed.Components;
using System.IO;

namespace InferMed.MACRO.ClinicalCoding.Plugins
{
	/// <summary>
	/// MedDRA preference object
	/// </summary>
	public class MedDRAPreference
	{
		//settings object
		IMEDSettings20 _iset = null;
		//browser columns
		private const string _COL_LLTKEY = "collltkey";
		private const string _COL_WEIGHT = "colweight";
		private const string _COL_FULLMATCH = "colfullmatch";
		private const string _COL_PARTMATCH = "colpartmatch";
		private const string _COL_PRIMARY = "colprimary";
		private const string _COL_CURRENT = "colcurrent";
		private const string _COL_PT = "colpt";
		private const string _COL_PTKEY = "colptkey";
		private const string _COL_HLT = "colhlt";
		private const string _COL_HLTKEY = "colhltkey";
		private const string _COL_HLGT = "colhlgt";
		private const string _COL_HLGTKEY = "colhlgtkey";
		private const string _COL_SOC = "colsoc";
		private const string _COL_SOCKEY = "colsockey";
		private const string _RESULT = "result";
		private const string _LEGEND = "legend";
		//preferences
		public bool _lltKey = true;
		public bool _weight = true;
		public bool _fullMatch = true;
		public bool _partMatch = true;
		public bool _primary = true;
		public bool _current = true;
		public bool _pt = true;
		public bool _ptKey = true;
		public bool _hlt = true;
		public bool _hltKey = true;
		public bool _hlgt = true;
		public bool _hlgtKey = true;
		public bool _soc = true;
		public bool _socKey = true;
		public int _result = 500;
		public bool _legend = false;
		//preference file
		private string _file = "";
		
		/// <summary>
		/// Load preferences from file into object
		/// </summary>
		/// <param name="file"></param>
		public MedDRAPreference(string file)
		{
			_file = file;
			if( !File.Exists( _file ) )
			{
				File.Create( _file );
			}
			else
			{
				_iset = new IMEDSettings20( _file );
				Load();
			}
		}

		/// <summary>
		/// Set preference from object
		/// </summary>
		private void Load()
		{
			_lltKey = System.Convert.ToBoolean( _iset.GetKeyValue( _COL_LLTKEY, "true" ) );
			_weight = System.Convert.ToBoolean( _iset.GetKeyValue( _COL_WEIGHT, "true" ) );
			_fullMatch = System.Convert.ToBoolean( _iset.GetKeyValue( _COL_FULLMATCH, "true" ) );
			_partMatch = System.Convert.ToBoolean( _iset.GetKeyValue( _COL_PARTMATCH, "true" ) );
			_primary = System.Convert.ToBoolean( _iset.GetKeyValue( _COL_PRIMARY, "true" ) );
			_current = System.Convert.ToBoolean( _iset.GetKeyValue( _COL_CURRENT, "true" ) );
			_pt = System.Convert.ToBoolean( _iset.GetKeyValue( _COL_PT, "true" ) );
			_ptKey = System.Convert.ToBoolean( _iset.GetKeyValue( _COL_PTKEY, "true" ) );
			_hlt = System.Convert.ToBoolean( _iset.GetKeyValue( _COL_HLT, "true" ) );
			_hltKey = System.Convert.ToBoolean( _iset.GetKeyValue( _COL_HLTKEY, "true" ) );
			_hlgt = System.Convert.ToBoolean( _iset.GetKeyValue( _COL_HLGT, "true" ) );
			_hlgtKey = System.Convert.ToBoolean( _iset.GetKeyValue( _COL_HLGTKEY, "true" ) );
			_soc = System.Convert.ToBoolean( _iset.GetKeyValue( _COL_SOC, "true" ) );
			_socKey = System.Convert.ToBoolean( _iset.GetKeyValue( _COL_SOCKEY, "true" ) );
			_result = System.Convert.ToInt32( _iset.GetKeyValue( _RESULT, "500" ) );
			_legend = System.Convert.ToBoolean( _iset.GetKeyValue( _LEGEND, "false" ) );
		}

		/// <summary>
		/// Save preferences to file
		/// </summary>
		public void Save()
		{
			if( _iset == null )
			{
				_iset = new IMEDSettings20( _file );
			}
			_iset.SetKeyValue( _COL_LLTKEY, System.Convert.ToString( _lltKey ) );
			_iset.SetKeyValue( _COL_WEIGHT, System.Convert.ToString( _weight ) );
			_iset.SetKeyValue( _COL_FULLMATCH, System.Convert.ToString( _fullMatch ) );
			_iset.SetKeyValue( _COL_PARTMATCH, System.Convert.ToString( _partMatch ) );
			_iset.SetKeyValue( _COL_PRIMARY, System.Convert.ToString( _primary ) );
			_iset.SetKeyValue( _COL_CURRENT, System.Convert.ToString( _current ) );
			_iset.SetKeyValue( _COL_PT, System.Convert.ToString( _pt ) );
			_iset.SetKeyValue( _COL_PTKEY, System.Convert.ToString( _ptKey ) );
			_iset.SetKeyValue( _COL_HLT, System.Convert.ToString( _hlt ) );
			_iset.SetKeyValue( _COL_HLTKEY, System.Convert.ToString( _hltKey ) );
			_iset.SetKeyValue( _COL_HLGT, System.Convert.ToString( _hlgt ) );
			_iset.SetKeyValue( _COL_HLGTKEY, System.Convert.ToString( _hlgtKey ) );
			_iset.SetKeyValue( _COL_SOC, System.Convert.ToString( _soc ) );
			_iset.SetKeyValue( _COL_SOCKEY, System.Convert.ToString( _socKey ) );
			_iset.SetKeyValue( _RESULT, System.Convert.ToString( _result ) );
			_iset.SetKeyValue( _LEGEND, System.Convert.ToString( _legend ) );
		}
	}
}
