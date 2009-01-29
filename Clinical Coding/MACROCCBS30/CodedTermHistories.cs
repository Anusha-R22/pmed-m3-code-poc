using System;
//using System.Runtime.InteropServices;
using System.Collections;
using System.Data;
using log4net.Config;
using log4net;
using System.IO;
using InferMed.Components;

namespace InferMed.MACRO.ClinicalCoding.MACROCCBS30
{
//	[ComVisible(true)]
//	[Guid("7A11AC95-5D3D-4E0D-80E7-ED23A44DE209")]
//	public interface ICodedTermHistories
//	{
//		[ComVisible(true)]
//		void InitAuto( string con, int clinicalTrialId, string trialSite, int personId, 
//			int crfPageId, short crfPageCycleNumber );
//		[ComVisible(true)]
//			void Save( string con, int visitId, short visitCycle, int crfPageId, short crfPageCycle );
//		[ComVisible(true)]
//			void AddCodedTermHistory( ref CodedTermHistory h );
//		[ComVisible(true)]
//			CodedTermHistory CodedTermHistoryFromTaskId( int responseTaskId, short repeat );
//		[ComVisible(true)]
//			bool Exists( int responseTaskId, short repeat );
//	}

	/// <summary>
	/// A collection of coded term current value and its historical values
	/// </summary>
//	[ComVisible(true)]
//	[Guid("8EB915E9-7D02-440C-81C9-7E292BB5090A")]
//	[ClassInterface(ClassInterfaceType.None)]
	public class CodedTermHistories //: ICodedTermHistories
	{
		//logging
		private static readonly ILog log = LogManager.GetLogger( typeof( CodedTermHistories ) );

		private ArrayList _codedTermHistories = null;

		public CodedTermHistories()
		{
			//use Init() as constructor because this is compiled for use in a vb6 application
			//and com does not support arguments in constructors
			//when this object is no longer being used by vb6, this constructor can be replaced 
			//and overloaded by methods below

			//initialise logging
			XmlConfigurator.Configure( new FileInfo( Path.GetDirectoryName( AppDomain.CurrentDomain.BaseDirectory ) 
				+ @"\log4netconfig.xml" ) );
			log.Info( "Initialised" );
		}

		/// <summary>
		/// Enforce integrity of clinical coding for a subject. 
		/// </summary>
		/// <param name="con"></param>
		/// <param name="clinicalTrialId"></param>
		/// <param name="trialSite"></param>
		/// <param name="personId"></param>
		/// <param name="checkOnly"></param>
		/// <param name="report"></param>
		/// <returns></returns>
		public bool EnforceIntegrity( string con, int clinicalTrialId, string trialSite, int personId )
		{
			DataSet dsDIR = null;
			DataSet dsDIRH = null;
			string sql = "";

			try
			{
				//get all current responses for thesaurus questions
				sql = "SELECT DIR.RESPONSETASKID, DIR.REPEATNUMBER, DIR.RESPONSETIMESTAMP, DIR.HADVALUE "
					+ "FROM DATAITEMRESPONSE DIR, DATAITEM DI "
					+ "WHERE DIR.CLINICALTRIALID = " + clinicalTrialId + " "
					+ "AND DIR.TRIALSITE = '" + trialSite + "' "
					+ "AND DIR.PERSONID = " + personId + " "
					+ "AND DIR.CLINICALTRIALID = DI.CLINICALTRIALID "
					+ "AND DIR.DATAITEMID = DI.DATAITEMID "
					+ "AND DI.DATATYPE = 8";

				dsDIR = CCDataAccess.GetDataSet( con, sql );

				//loop through all the clinical coding responses for this person
				foreach (DataRow rowDIR in dsDIR.Tables[0].Rows)
				{
					//get the coded term history for this response
					CodedTermHistory cth = new CodedTermHistory();
					cth.InitAuto( con, clinicalTrialId, trialSite, personId, Convert.ToInt32(rowDIR["RESPONSETASKID"].ToString()), 
						Convert.ToInt16(rowDIR["REPEATNUMBER"].ToString()));

					if ((!cth.NoHistory) && (cth.ResponseTimeStamp == Convert.ToDouble(CCDataAccess.ConvertFromNull(rowDIR, "RESPONSETIMESTAMP"))))
					{
						//the response timestamp of the current response value matches the response value saved
						//against the coding history record - this response hasnt been changed since it was coded
						//-no need to do anything
					}
					else
					{
						if ((!cth.NoHistory) && ((cth.CodingStatus == CodedTerm.eCodingStatus.DoNotCode) 
							|| (cth.CodingStatus == CodedTerm.eCodingStatus.NotCoded)))
						{
							//the current coding status is 'do not code' or 'not coded' - we dont need to add any coded term
							//history records, but need to update the coding status of the response in the DIR table
							CCDataAccess.RunSQL( con, CCDataAccess.Sproc_SP_MACRO_CODING_UPDATE_DIR( IMEDDataAccess.CalculateConnectionType( con ), 
								clinicalTrialId , trialSite, personId, cth.ResponseTaskId, cth.Repeat, cth.DictionaryName, 
								cth.DictionaryVersion, cth.CodingStatus, cth.CodingDetails));
						}
						else
						{
							//(a) the current coding status is one which must revert to 'not coded' after a response change -
							//add a coded term history record including the value that initially caused the 
							//change to 'not coded'
							//(b) no coded term history exists - create an initial history of 'not coded' if there has
							//ever been a response value

							//get the responsetimestamp of the response where
							//(a) the value changed from the codedhistory value, and the responsetimestamp was later 
							//than the coded history value
							//(b) the value was no longer ""
							sql = "SELECT MIN(RESPONSETIMESTAMP) AS RESPONSETIMESTAMP FROM DATAITEMRESPONSEHISTORY "
								+ "WHERE CLINICALTRIALID = " + clinicalTrialId + " "
								+ "AND TRIALSITE = '" + trialSite + "' "
								+ "AND PERSONID = " + personId + " "
								+ "AND RESPONSETASKID = " + rowDIR["RESPONSETASKID"].ToString() + " "
								+ "AND REPEATNUMBER = " + rowDIR["REPEATNUMBER"].ToString() + " "
								+ "AND RESPONSEVALUE <> '" + ((cth.NoHistory) ? "" : cth.ResponseValue) + "'"
								+ ((cth.NoHistory) ? "" : " AND RESPONSETIMESTAMP > " + cth.ResponseTimeStamp.ToString());

							dsDIRH = CCDataAccess.GetDataSet( con, sql );
							string responseTimeStamp = CCDataAccess.ConvertFromNull(dsDIRH.Tables[0].Rows[0], "RESPONSETIMESTAMP");

							if (responseTimeStamp == "")
							{
								//there IS NOT a response where
								//(a) the value is different and the timestamp later than the coded history
								//(b) the value <> ""
								if (!cth.NoHistory)
								{
									//(a)
									//need to update the response row in the DIR table, no changes to the codinghistory
									CCDataAccess.RunSQL( con, CCDataAccess.Sproc_SP_MACRO_CODING_UPDATE_DIR( IMEDDataAccess.CalculateConnectionType( con ), 
										clinicalTrialId , trialSite, personId, cth.ResponseTaskId, cth.Repeat, cth.DictionaryName, 
										cth.DictionaryVersion, cth.CodingStatus, cth.CodingDetails));
								}
								else
								{
									//(b)
									//dont need to a do anything - the response has never had a value
								}
							}
							else
							{
								//there IS a response where
								//(a) the value is different and the timestamp later than the coded history
								//(b) the value <> ""
									
								//get the DIRH row for this timestamp
								sql = "SELECT CRFPAGEID, CRFPAGECYCLENUMBER, VISITID, VISITCYCLENUMBER, RESPONSEVALUE, RESPONSETIMESTAMP, "
									+ "RESPONSETIMESTAMP_TZ, USERNAME, USERNAMEFULL FROM DATAITEMRESPONSEHISTORY "
									+ "WHERE CLINICALTRIALID = " + clinicalTrialId + " "
									+ "AND TRIALSITE = '" + trialSite + "' "
									+ "AND PERSONID = " + personId + " "
									+ "AND RESPONSETASKID = " + rowDIR["RESPONSETASKID"].ToString() + " "
									+ "AND REPEATNUMBER = " + rowDIR["REPEATNUMBER"].ToString() + " "
									+ "AND RESPONSETIMESTAMP = " + responseTimeStamp;

								dsDIRH = CCDataAccess.GetDataSet( con, sql );
								DataRow rowDIRH = dsDIRH.Tables[0].Rows[0];
							
								//create the coding status and the DIR updates
								cth.SetStatus(CodedTerm.eCodingStatus.NotCoded, rowDIRH["USERNAME"].ToString(), rowDIRH["USERNAMEFULL"].ToString(),
									rowDIRH["RESPONSEVALUE"].ToString(), Convert.ToDouble(rowDIRH["RESPONSETIMESTAMP"].ToString()),
									Convert.ToInt16(rowDIRH["RESPONSETIMESTAMP_TZ"].ToString()));

								//commit the changes
								cth.Save(con, Convert.ToInt32(rowDIRH["VISITID"].ToString()), Convert.ToInt16(rowDIRH["VISITCYCLENUMBER"].ToString()),
									Convert.ToInt32(rowDIRH["CRFPAGEID"].ToString()), Convert.ToInt16(rowDIRH["CRFPAGECYCLENUMBER"].ToString()));
							}
						}
					}
				}
				return true;
			}
#if !DEBUG
			catch
			{
				return false;
			}
#endif
			finally
			{
				dsDIR.Dispose();
				dsDIRH.Dispose();
			}
		}

		/// <summary>
		/// Automatic initialisation
		/// </summary>
		/// <param name="con"></param>
		/// <param name="clinicalTrialId"></param>
		/// <param name="trialSite"></param>
		/// <param name="personId"></param>
		/// <param name="crfPageId"></param>
		/// <param name="crfPageCycleNumber"></param>
//		[ComVisible(true)]
		public void InitAuto( string con, int clinicalTrialId, string trialSite, int personId, int crfPageId, 
			short crfPageCycleNumber )
		{
			DataSet ds = null;
			int responseTaskId = 0;
			short repeat = 0;
			
			try
			{
				//create sql all coding histories for this eform
				string sql = "SELECT * FROM CODINGHISTORY "
					+ "WHERE CLINICALTRIALID = " + clinicalTrialId + " "
					+ "AND TRIALSITE = '" + trialSite + "' " 
					+ "AND PERSONID = " + personId + " "
					+ "AND CRFPAGEID = " + crfPageId + " "
					+ "AND CRFPAGECYCLENUMBER = " + crfPageCycleNumber + " "
					+ "ORDER BY RESPONSETASKID, REPEATNUMBER ASC";

				//get the coding histories, if any
				ds = CCDataAccess.GetDataSet( con, sql );

				CodedTermHistory h = null;
				//add each history row to the object as a new codedterm
				foreach( DataRow r in ds.Tables[0].Rows )
				{
					if( ( System.Convert.ToInt32( CCDataAccess.ConvertFromNull( r, "ResponseTaskId" ) ) != responseTaskId ) 
						|| ( System.Convert.ToInt16( CCDataAccess.ConvertFromNull( r, "RepeatNumber" ) ) != repeat ) )
					{
						//if this is the first of a new response, set the current response params and create a new object
						h = new CodedTermHistory();
						AddCodedTermHistory( ref h );
						responseTaskId = System.Convert.ToInt32( CCDataAccess.ConvertFromNull( r, "ResponseTaskId" ) );
						repeat = System.Convert.ToInt16( CCDataAccess.ConvertFromNull( r, "RepeatNumber" ) );
					}

					if( CodedTerm.GetStatus( CCDataAccess.ConvertFromNull( r, "Status" ) ) == CodedTerm.eStatus.Current )
					{
						//if this is the current coded value, set the current object properties
						h.InitManual( clinicalTrialId, trialSite, personId, responseTaskId, repeat, 
							CCDataAccess.ConvertFromNull( r, "DictionaryName" ),	
							CCDataAccess.ConvertFromNull( r, "DictionaryVersion" ), 
							CodedTerm.GetCodingStatus( CCDataAccess.ConvertFromNull( r, "CodingStatus" ) ),
							CCDataAccess.ConvertFromNull( r, "CodingDetails" ), 
							System.Convert.ToDouble( CCDataAccess.ConvertFromNull( r, "CodingTimeStamp" ) ),	
							System.Convert.ToInt16( CCDataAccess.ConvertFromNull( r, "CodingTimeStamp_TZ" ) ),
							CCDataAccess.ConvertFromNull( r, "ResponseValue" ),
							System.Convert.ToDouble( CCDataAccess.ConvertFromNull( r, "ResponseTimeStamp" ) ), 
							System.Convert.ToInt16( CCDataAccess.ConvertFromNull( r, "ResponseTimeStamp_TZ" ) ), 
							CCDataAccess.ConvertFromNull( r, "UserName" ),
							CCDataAccess.ConvertFromNull( r, "UserNameFull" ) );
					}

					//add the coded value to the history list
					CodedTerm t = new CodedTerm();

					t.Init( CCDataAccess.ConvertFromNull( r, "DictionaryName" ),	
						CCDataAccess.ConvertFromNull( r, "DictionaryVersion" ), 
						CodedTerm.GetCodingStatus( CCDataAccess.ConvertFromNull( r, "CodingStatus" ) ),
						CCDataAccess.ConvertFromNull( r, "CodingDetails" ), 
						System.Convert.ToDouble( CCDataAccess.ConvertFromNull( r, "CodingTimeStamp" ) ),	
						System.Convert.ToInt16( CCDataAccess.ConvertFromNull( r, "CodingTimeStamp_TZ" ) ),
						CCDataAccess.ConvertFromNull( r, "ResponseValue" ),
						System.Convert.ToDouble( CCDataAccess.ConvertFromNull( r, "ResponseTimeStamp" ) ), 
						System.Convert.ToInt16( CCDataAccess.ConvertFromNull( r, "ResponseTimeStamp_TZ" ) ), 
						CCDataAccess.ConvertFromNull( r, "UserName" ),
						CCDataAccess.ConvertFromNull( r, "UserNameFull" ), 
						CCDataAccess.ConvertFromNull( r, "ReasonForChange" ) );
					
					h.AddHistoryTerm( ref t );
				}
			}
			finally
			{
				ds.Dispose();
			}
		}

		/// <summary>
		/// Save all coded terms on an eform
		/// </summary>
		/// <param name="con"></param>
		/// <param name="visitId"></param>
		/// <param name="visitCycle"></param>
		/// <param name="crfPageId"></param>
		/// <param name="crfPageCycle"></param>
//		[ComVisible(true)]
		public void Save( string con, int visitId, short visitCycle, int crfPageId, short crfPageCycle )
		{
			if( _codedTermHistories != null )
			{
				//loop through all coded terms in the list, saving 
				for( int n = 0; n < _codedTermHistories.Count; n++ )
				{
					CodedTermHistory h = ( CodedTermHistory ) _codedTermHistories[n];
					h.Save( con, visitId, visitCycle, crfPageId, crfPageCycle );
				}
			}
		}

		/// <summary>
		/// Add a codedterm to the history list
		/// </summary>
		/// <param name="h"></param>
//		[ComVisible(true)]
		public void AddCodedTermHistory( ref CodedTermHistory h )
		{
			if( _codedTermHistories == null ) _codedTermHistories = new ArrayList();
			_codedTermHistories.Add( h );
		}

		/// <summary>
		/// Get a codedterm by its taskid and repeat
		/// </summary>
		/// <param name="responseTaskId"></param>
		/// <param name="repeat"></param>
		/// <returns></returns>
//		[ComVisible(true)]
		public CodedTermHistory CodedTermHistoryFromTaskId( int responseTaskId, short repeat )
		{
			CodedTermHistory h = null;
			int n = 0;

			if( _codedTermHistories != null )
			{
				while( ( h == null ) && ( n < _codedTermHistories.Count ) )
				{
					CodedTermHistory temph = ( CodedTermHistory )_codedTermHistories[n];
					if( ( temph.ResponseTaskId == responseTaskId ) 
						&& ( temph.Repeat == repeat ) ) h = temph;
					n++;
				}
			}
			return( h );
		}

		/// <summary>
		/// Does a codedterm exist in the list
		/// </summary>
		/// <param name="responseTaskId"></param>
		/// <param name="repeat"></param>
		/// <returns></returns>
//		[ComVisible(true)]
		public bool Exists( int responseTaskId, short repeat )
		{
			if( _codedTermHistories != null )
			{
				foreach( object o in _codedTermHistories )
				{
					CodedTermHistory h = ( CodedTermHistory )o;
					if( ( h.ResponseTaskId == responseTaskId )
						&& ( h.Repeat == repeat ) ) return( true );
				}
			}
			return( false );
		}

		/// <summary>
		/// Have any of the values of any of the codedterms changed requiring a save
		/// </summary>
		public bool SaveNeeded
		{
			get
			{
				if( _codedTermHistories != null )
				{
					for( int n = 0; n < _codedTermHistories.Count; n++ )
					{
						CodedTermHistory h = ( CodedTermHistory ) _codedTermHistories[n];
						if( h.IsEdited ) return( true );
					}
				}
				return( false );
			}
		}

		/// <summary>
		/// How many codedterms are in the list
		/// </summary>
		public int Count
		{
			get { return( ( _codedTermHistories == null ) ? 0 :  _codedTermHistories.Count ); }
		}

		/// <summary>
		/// Get the list of codedterms
		/// </summary>
//		[ComVisible(true)]
		public ArrayList CodedTermHistoryList
		{
			get 
			{ 
				if( _codedTermHistories == null ) _codedTermHistories = new ArrayList();
				return( _codedTermHistories ); 
			}
		}
	}
}
