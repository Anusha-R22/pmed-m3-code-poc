using System;
using System.Collections;

namespace InferMed.MACRO.StudyCopy
{
	/// <summary>
	/// MACRO eform object
	/// </summary>
	public class Eform
	{
		//element list
		private ArrayList _elements = new ArrayList();

		//eform properties
		private string _destinationId = "";
		private string _sourceId = "";
		private bool _copied = false;
//		private bool _matched = false;
		private string _copyDate = "";

		/// <summary>
		/// Unmatch this eform and its elements
		/// </summary>
		public void UnMatch()
		{
			_sourceId = "";
			_copied = false;
//			_matched = false;
			_copyDate = "";
			_elements.Clear();
		}

		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="id"></param>
		public Eform( string DestinationId )
		{
			_destinationId = DestinationId;
		}

		/// <summary>
		/// Eform id
		/// </summary>
		public string DestinationId
		{
			get{ return( _destinationId ); }
		}

		/// <summary>
		/// Matched source eform id
		/// </summary>
		public string SourceId
		{
			get{ return( _sourceId ); }
			set{ _sourceId = value; }
		}

		/// <summary>
		/// Has the eform been copied
		/// </summary>
		public bool Copied
		{
			get{ return( _copied ); }
			set
			{ 
				_copyDate = DateTime.Now.ToString();
				_copied = value; 
			}
		}

		/// <summary>
		/// Add an eform element to the element list
		/// </summary>
		/// <param name="element"></param>
		public void AddElement( EformElement element )
		{
			if( GetElementByDestination( element.DestinationId ) == null )
			{
				_elements.Add( element );
			}
		}

		/// <summary>
		/// Get an eform element from the element list, using the destination element id
		/// </summary>
		/// <param name="id"></param>
		/// <returns></returns>
		public EformElement GetElementByDestination( string destinationId )
		{
			foreach( EformElement el in _elements )
			{
				if( el.DestinationId == destinationId ) return( el );
			}

			return( null );
		}

		/// <summary>
		/// Get an eform element from the element list, using the source element id
		/// </summary>
		/// <param name="sourceId"></param>
		/// <returns></returns>
		public EformElement GetElementBySource( string sourceId )
		{
			foreach( EformElement el in _elements )
			{
				if( el.SourceId == sourceId ) return( el );
			}

			return( null );
		}

		/// <summary>
		/// Elements
		/// </summary>
		public ArrayList Elements
		{
			get{ return( _elements ); }
		}

		/// <summary>
		/// Has the eform been matched up
		/// </summary>
		public bool Matched
		{
			get{ return( _sourceId != "" ); }
		}

		/// <summary>
		/// Are all elements of the eform copied
		/// </summary>
		public bool AllElementsCopied
		{
			get
			{
				foreach( EformElement el in _elements )
				{
					if( !el.Copied ) return( false );
				}
				return( true );
			}
		}
	}
}
