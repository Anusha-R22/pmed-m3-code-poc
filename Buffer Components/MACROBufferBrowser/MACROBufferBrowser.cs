using System;
using log4net;
using log4net.Config;
using System.Runtime.InteropServices;
using System.IO;

namespace InferMed.MACROBuffer
{
	/// <summary>
	/// Main Buffer Data Browser class
	/// </summary>
	[ComVisible(true)]
	[Guid("71174b06-be43-40b6-95db-c756893ccfcc")]
	[ClassInterface(ClassInterfaceType.None)]
	public class MACROBufferBrowser : IMACROBufferBrowser
	{
		// log4net
		private static readonly ILog log = LogManager.GetLogger( typeof(MACROBufferBrowser) );

		public MACROBufferBrowser()
		{
			// Need to leave empty as using with COM
		}
		
		/// <summary>
		/// Init function - included as using with COM
		/// </summary>
		public void Init()
		{
			// log4net initialisation
			XmlConfigurator.Configure( new System.IO.FileInfo( Path.GetDirectoryName( AppDomain.CurrentDomain.BaseDirectory ) + @"\log4netconfig.xml") );
		}

		/// <summary>
		/// Buffer Summary page
		/// </summary>
		/// <param name="serialisedUser">serialised user string</param>
		/// <param name="isUserHex">is serialised user hex</param>
		/// <param name="studyId">study id</param>
		/// <param name="site">site code</param>
		/// <param name="subjectNo">subject no</param>
		/// <returns>Summary page html</returns>
		[ComVisible(true)]
		public string BufferSummaryPage(string serialisedUser, bool isUserHex, int studyId, string site, int subjectNo)
		{
			// initialise component & log4net
			Init();

			// create user object
			BufferMACROUser bufferUser;

			// buffer page HTML to return
			string pageHTML = "";

			try
			{
				log.Info("starting BufferSummaryPage");
				log.Info("serialised user length=" + serialisedUser.Length);
				log.Info("Working directory=" + Path.GetDirectoryName( AppDomain.CurrentDomain.BaseDirectory ));
				log.Info("Study=" + studyId.ToString() + " Site=" + site.ToString() + " Subject=" + subjectNo.ToString() );

				// create user object
				bufferUser = new BufferMACROUser(serialisedUser, isUserHex);

				// render summary page 
				pageHTML = BufferSummary.RenderBufferPage( bufferUser, studyId, site, subjectNo );

				// dispose of user object
				bufferUser.Dispose();
			}
			catch(Exception ex)
			{
				log.Error(ex.Message);
			}
			return pageHTML;
		}

		/// <summary>
		/// Load Buffer Data Browser
		/// </summary>
		/// <param name="serialisedUser">serialised user string</param>
		/// <param name="isUserHex">is serialised user hex</param>
		/// <param name="studyId">study id</param>
		/// <param name="site">site code</param>
		/// <param name="subjectNo">subject no</param>
		/// <param name="bookMark">bookmark value from which to render page</param>
		/// <returns>main buffer data browser html page</returns>
		[ComVisible(true)]
		public string LoadBufferDataBrowser(string serialisedUser, bool isUserHex, int studyId, string site, int subjectNo,
								string bookMark)
		{
			// initialise component & log4net
			Init();

			// create user object
			BufferMACROUser bufferUser;

			// buffer browser page HTML to return
			string pageHTML = "";

			try
			{
				log.Info("starting LoadBufferDataBrowser");
				log.Info("Working directory=" + Path.GetDirectoryName( AppDomain.CurrentDomain.BaseDirectory ));
				log.Info("Study=" + studyId.ToString() + " Site=" + site.ToString() + " Subject=" + subjectNo.ToString() + " bookMark=" + bookMark);

				// create user object
				bufferUser = new BufferMACROUser(serialisedUser, isUserHex);

				// create buffer browser object
				BufferDataBrowser bufferDataBrowser = new BufferDataBrowser(bufferUser, studyId, site, subjectNo, bookMark);

				// render buffer browser page 
				pageHTML = bufferDataBrowser.RenderBufferBrowserPage( bufferUser, studyId, site, subjectNo );

				// dispose of user object
				bufferUser.Dispose();
			}
			catch(Exception ex)
			{
				log.Error(ex.Message);
			}
			return pageHTML;
		}

		/// <summary>
		/// Get buffer save results
		/// </summary>
		/// <param name="serialisedUser">serialised user string</param>
		/// <param name="isUserHex">is serialised user hex</param>
		/// <param name="formData">asp pages form data</param>
		/// <returns>Buffer Save results page html</returns>
		[ComVisible(true)]
		public string GetBufferSaveResultsPage(string serialisedUser, bool isUserHex, string formData)
		{
			// initialise component & log4net
			Init();

			// create user object
			BufferMACROUser bufferUser;

			// buffer data save results page HTML to return
			string pageHTML = "";

			try
			{
				log.Info("starting GetBufferSaveResultsPage");

				// create user object
				bufferUser = new BufferMACROUser(serialisedUser, isUserHex);

				// create buffer save results object
				BufferDataBrowserSave bufferDataBrowserSave = new BufferDataBrowserSave( formData );

				// render buffer save results page 
				pageHTML = bufferDataBrowserSave.RenderBufferSaveResultsPage( bufferUser );

				// dispose of user object
				bufferUser.Dispose();
			}
			catch(Exception ex)
			{
				log.Error(ex);
			}

			return pageHTML;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="serialisedUser">serialised user string</param>
		/// <param name="isUserHex">is serialised user hex</param>
		/// <param name="formData">asp pages form data</param>
		/// <param name="bookMark"></param>
		/// <returns>buffer target selection page html</returns>
		[ComVisible(true)]
		public string GetBufferTargetSelectionPage(string serialisedUser, bool isUserHex, string formData, string bookMark)
		{
			// initialise component & log4net
			Init();

			// create user object
			BufferMACROUser bufferUser;

			// buffer target selection page HTML to return
			string pageHTML = "";

			try
			{
				log.Info("starting GetBufferTargetSelectionPage");

				// create user object
				bufferUser = new BufferMACROUser(serialisedUser, isUserHex);

				// create buffer save results object
				BufferTargetSelection bufferTargetSelection = new BufferTargetSelection( formData, bookMark );

				// render buffer save results page 
				pageHTML = bufferTargetSelection.RenderBufferTargetPage( bufferUser );

				// dispose of user object
				bufferUser.Dispose();
			}
			catch(Exception ex)
			{
				log.Error(ex);
			}

			return pageHTML;
		}

		/// <summary>
		/// this function will save the buffer data and automatically return to the buffer browser page
		/// </summary>
		/// <param name="serialisedUser">serialised user string</param>
		/// <param name="isUserHex">is serialised user hex</param>
		/// <param name="formData">asp pages form data</param>
		/// <returns>buffer target selection save page html</returns>
		[ComVisible(true)]
		public string SaveBufferTargetSelection( string serialisedUser, bool isUserHex, string formData )
		{
			// initialise component & log4net
			Init();

			// create user object
			BufferMACROUser bufferUser;

			// buffer target selection page HTML to return
			string pageHTML = "";

			try
			{
				log.Info("starting SaveBufferTargetSelection");

				// create user object
				bufferUser = new BufferMACROUser(serialisedUser, isUserHex);

				// create buffer save results object
				BufferTargetSelectionSave bufferTargetSelectionSave = new BufferTargetSelectionSave( formData );

				// render buffer save results page 
				// this page will save the buffer data and automatically return to the buffer browser page
				pageHTML = bufferTargetSelectionSave.SaveBufferTarget( bufferUser );

				// dispose of user object
				bufferUser.Dispose();
			}
			catch(Exception ex)
			{
				log.Error(ex);
			}

			return pageHTML;
		}

		/// <summary>
		/// return working directory
		/// </summary>
		/// <returns></returns>
		[ComVisible(true)]
		public string WorkingDirectory()
		{
			log.Info("starting WorkingDirectory");

			return Path.GetDirectoryName( AppDomain.CurrentDomain.BaseDirectory );
		}
	}
}
