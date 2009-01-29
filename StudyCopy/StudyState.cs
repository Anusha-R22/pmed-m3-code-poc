using System;
using System.Collections;

namespace InferMed.MACRO.StudyCopy
{
	/// <summary>
	/// Study state object
	/// </summary>
	public class StudyState
	{
		//eform list
		private ArrayList _eforms = new ArrayList();

		//copied dataitem list
		private ArrayList _dataItems = new ArrayList();

		//copied question group list
		private ArrayList _questionGroups = new ArrayList();

		//state file, if any
		private string _stateFile = "";

		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="stateFile"></param>
		public StudyState( string stateFile )
		{
			_stateFile = stateFile;
		}

		/// <summary>
		/// Unmatch all destination eforms
		/// </summary>
		public void Unmatch()
		{
			foreach( Eform ef in _eforms )
			{
				ef.UnMatch();
				_dataItems.Clear();
				_questionGroups.Clear();
			}
		}

		/// <summary>
		/// Clear the state object
		/// </summary>
		public void Clear()
		{
			_eforms.Clear();
			_dataItems.Clear();
			_questionGroups.Clear();
		}

		/// <summary>
		/// Is a dataitem in the list (if not add it to the list)
		/// </summary>
		/// <param name="eformId"></param>
		/// <returns></returns>
		public bool DataItemCopied( string dataItemId )
		{
			if( ( dataItemId != "0" ) && ( dataItemId != "" ) )
			{
				foreach( string id in _dataItems )
				{
					if( id == dataItemId ) return( true );
				}
				_dataItems.Add( dataItemId );
				return( false );
			}
			else
			{
				return( true );
			}
			
		}

		/// <summary>
		/// Is a qgroup in the list (if not add it to the list)
		/// </summary>
		/// <param name="eformId"></param>
		/// <returns></returns>
		public bool QGroupCopied( string qGroupId )
		{
			if( ( qGroupId != "0" ) && ( qGroupId != "" ) )
			{	
				foreach( string id in _questionGroups )
				{
					if( id == qGroupId ) return( true );
				}																				
				_questionGroups.Add( qGroupId );
				return( false );
			}
			else
			{
				return( true );
			}
			
		}

		/// <summary>
		/// Add an eform element to the eform list
		/// </summary>
		/// <param name="ef"></param>
		public void AddEform( Eform ef )
		{
			if( GetEform( ef.DestinationId ) == null )
			{
				_eforms.Add( ef );
			}
		}

		/// <summary>
		/// Get an eform from the eform list
		/// </summary>
		/// <param name="eformId"></param>
		/// <returns></returns>
		public Eform GetEform( string eformId )
		{
			foreach( Eform ef in _eforms )
			{
				if( ef.DestinationId == eformId ) return( ef );
			}

			return( null );
		}

		/// <summary>
		/// Eforms
		/// </summary>
		public ArrayList Eforms
		{
			get{ return( _eforms ); }
		}
	}
}
