using System;
using MACROUserBS30;
using log4net;

namespace InferMed.MACROBuffer
{
	/// <summary>
	/// Holds the MACRO user object
	/// </summary>
	class BufferMACROUser : IDisposable
	{
		// user object
		private MACROUserBS30.MACROUserClass _MACROUser;
		// has class been disposed?
		private bool _isDisposed=false;
		// log4net
		private static readonly ILog log = LogManager.GetLogger( typeof(BufferMACROUser) );

		/// <summary>
		/// constructor for creating MACRO user object
		/// </summary>
		/// <param name="serialisedUser"></param>
		/// <param name="bHex"></param>
		public BufferMACROUser(string serialisedUser, bool bHex)
		{
			try
			{
				// create new user object
				_MACROUser = new MACROUserClass();
				log.Debug("Set the MACRO user state - hex=" + bHex.ToString());
				if(bHex)
				{
					// set the state (hex)
					_MACROUser.SetStateHex(ref serialisedUser);
				}
				else
				{
					// set the state 
					_MACROUser.SetState(ref serialisedUser);
				}
			}
			catch(Exception ex)
			{
				// log
				log.Error( "Error initialising MACRO user object", ex );
				// rethrow
				throw (new Exception(ex.Message));
			}
		}
		// properties
		/// <summary>
		/// Allow access to the MACRO User object through this property
		/// </summary>
		public MACROUserClass MACROUser
		{
			get
			{
				return _MACROUser;
			}
		}

		/// <summary>
		/// public dispose of the MACRO user object
		/// </summary>
		public void Dispose()
		{
			// dispose of user object
			Dispose(true);
			// tell garbage collector not to worry about this class now 
			GC.SuppressFinalize(this);
		}

		/// <summary>
		/// private dispose of the MACRO user object
		/// </summary>
		/// <param name="bDispose">dispose of managed objects</param>
		protected void Dispose(bool bDispose)
		{
			// check if already disposed
			if(!_isDisposed)
			{
				// check if called from code
				if(bDispose)
				{
					// remove user object
					_MACROUser = null;
				}
				// set private flag to having been disposed
				_isDisposed=true;
			}
		}

		/// <summary>
		/// Destructor
		/// </summary>
		~BufferMACROUser()
		{
			Dispose(false);
		}
	}
}
