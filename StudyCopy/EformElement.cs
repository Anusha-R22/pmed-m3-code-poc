using System;

namespace InferMed.MACRO.StudyCopy
{
	/// <summary>
	/// A MACRO eform element
	/// </summary>
	public class EformElement
	{
		//element properties
		private string _destinationId;
		private string _sourceId = "";
		private bool _copied = false;
		private string _copyDate = "";


		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="id"></param>
		public EformElement( string destinationId )
		{
			_destinationId = destinationId;
		}

		/// <summary>
		/// Element id
		/// </summary>
		public string DestinationId
		{
			get{ return( _destinationId ); }
		}

		/// <summary>
		/// Matched source element id
		/// </summary>
		public string SourceId
		{
			get{ return( _sourceId ); }
			set{ _sourceId = value; }
		}

		/// <summary>
		/// Has the element been copied
		/// </summary>
		public bool Copied
		{
			get{ return( _copied ); }
			set
			{
				_copied = value;
				_copyDate = DateTime.Now.ToString();
			}
		}

		/// <summary>
		/// Has the element been matched up
		/// </summary>
		public bool Matched
		{
			get{ return( _sourceId != "" ); }
		}

		/// <summary>
		/// Unmatch the element
		/// </summary>
		public void UnMatch()
		{
			_sourceId = "";
			_copied = false;
			_copyDate = "";
		}
	}
}
