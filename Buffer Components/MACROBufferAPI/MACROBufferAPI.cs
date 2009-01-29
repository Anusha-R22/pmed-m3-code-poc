using System;
using System.IO;
using log4net;
using log4net.Config;
using System.Runtime.InteropServices;

namespace InferMed.MACROBuffer
{
	/// <summary>
	/// Summary description for Class1.
	/// </summary>
	[ComVisible(true)]
	[Guid("9bb110b9-232b-4aeb-81f6-7c96f6e66a22")]
	[ClassInterface(ClassInterfaceType.None)]
	public class MACROBufferAPI : IMACROBufferAPI
	{
		// member variables
		private SubjectDetails _subjectDetail = null;
		
		// log4net
		private static readonly ILog log = LogManager.GetLogger( typeof(MACROBufferAPI) );

		public MACROBufferAPI()
		{
			// Need to leave empty as using with COM
		}

		// Init function - included as using with COM
		public void Init()
		{
			// new subject detail
			_subjectDetail = new SubjectDetails();
			// log4net initialisation
			XmlConfigurator.Configure( new System.IO.FileInfo( Path.GetDirectoryName( AppDomain.CurrentDomain.BaseDirectory ) + @"\log4netconfig.xml") );
		}

		/// <summary>
		/// Main method - attempts to write buffer message to the MACRO database - COM visible
		/// </summary>
		/// <param name="sBufferMessageXML"></param>
		/// <param name="bCommitToMACRO"></param>
		/// <returns></returns>
		[ComVisible(true)]
		public string WriteBufferMessage(string sBufferMessageXML, bool bCommitToMACRO)
		{
			// create subjectdetails object
			Init();

			string sMessageReportXML = "";

			// parse message xml
			BufferMessageXML bufferXML = new BufferMessageXML();
			bool bMACROCommitOK = true;
			// log
			log.Info("Parse buffer message.");
			_subjectDetail.SubjectResponseStatus = bufferXML.ParseBufferMessage( sBufferMessageXML, ref _subjectDetail );
			if(_subjectDetail.SubjectResponseStatus != BufferAPI.BufferResponseStatus.Success)
			{
				bMACROCommitOK = false;
			}

			// create database connection
			BufferResponseDb bufferDb = null;
			bool bDbConnOk = true;

			try
			{
				// log
				log.Info("Create database connection.");
				bufferDb = new BufferResponseDb();
			}
			catch(Exception ex)
			{
				// database error
				_subjectDetail.SubjectResponseStatus = BufferAPI.BufferResponseStatus.FailDbConnection;
				// log
				log.Error("Error creating database connection", ex);
				bDbConnOk = false;
				bMACROCommitOK = false;
			}

			if(bDbConnOk)
			{
				// check if buffer API tables exist
				if( ! bufferDb.DoBufferAPITablesExist() )
				{
					// create tables
					log.Info("Create Buffer API tables.");
					try
					{
						bufferDb.CreateBufferAPITables();
					}
					catch(Exception ex)
					{
						log.Error("Error creating buffer api database tables", ex);
						bDbConnOk = false;
						bMACROCommitOK = false;
					}
				}

				// if passed parse
				if(_subjectDetail.SubjectResponseStatus == BufferAPI.BufferResponseStatus.Success)
				{
					// perform basic checks
					log.Info("Perform basic checks");
					if(!BufferBasicChecks.PerformBasicChecks( ref _subjectDetail, bufferDb ))
					{
						bMACROCommitOK = false;
					}
				}

				// store message to BufferResponse table
				try
				{
					log.Info("Write buffer message to database.");
					bufferDb.WriteBufferMessages( ref sBufferMessageXML, ref _subjectDetail);
				}
				catch
				{
					// database error
					_subjectDetail.SubjectResponseStatus = BufferAPI.BufferResponseStatus.FailDbConnection;
					bDbConnOk = false;
					bMACROCommitOK = false;
				}

				// commit to MACRO directly?
				if((bCommitToMACRO)&&(bDbConnOk)&&(bMACROCommitOK))
				{
					// attempt write to MACRO database
					// require ALL fields complete to do this
					// need to pass _subjectDetail, bufferDb
					try
					{
						log.Info("Attempt direct commit of buffer rows to MACRO response table.");
						MACROCommit macroCommit = new MACROCommit(BufferAPI.GetSetting(BufferAPI._MACRO_CONN_DB,""));
						macroCommit.WriteToMACRO(ref _subjectDetail, bufferDb);
					}
					catch(Exception ex)
					{
						// log error
						log.Error("Error committing data directly to MACRO Db", ex);
					}
				}
			}

			// format message report xml
			log.Info("Create message XML report.");
			sMessageReportXML = bufferXML.CreateMessageXMLReport( ref _subjectDetail, ((bCommitToMACRO)&&(bDbConnOk)&&(bMACROCommitOK)) );

			return sMessageReportXML;
		}
	}
}
