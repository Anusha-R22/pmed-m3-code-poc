using System;
using System.Text;
using log4net;

namespace InferMed.MACROBuffer
{
	/// <summary>
	/// Summary description for BufferBasicChecks.
	/// </summary>
	class BufferBasicChecks
	{
		// log4net
		private static readonly ILog log = LogManager.GetLogger( typeof(BufferAPI) );

		private BufferBasicChecks()
		{}

		public static bool PerformBasicChecks(ref SubjectDetails subjDetail, BufferResponseDb bufferDb)
		{
			bool bOk=true;
			// check study definition
			try
			{
				subjDetail.ClinicalTrialId = bufferDb.GetStudyDefinitionId(subjDetail.ClinicalTrialName);
			}
			catch(Exception ex)
			{
				// study definition doesn't exist
				subjDetail.SubjectResponseStatus = BufferAPI.BufferResponseStatus.FailClinicalTrial;
				// log error
				StringBuilder sbError = new StringBuilder();
				sbError.Append( "Study with name '" );
				sbError.Append( subjDetail.ClinicalTrialName );
				sbError.Append( "' does not exist on the database." );
				log.Error( sbError.ToString(), ex );
				return false;
			}

			// check study / site
			try
			{
				bufferDb.CheckStudySite( subjDetail.ClinicalTrialId, subjDetail.Site );
			}
			catch(Exception ex)
			{
				// study / site combo do not exist
				subjDetail.SubjectResponseStatus = BufferAPI.BufferResponseStatus.FailSite;
				// log error
				StringBuilder sbError = new StringBuilder();
				sbError.Append( "Study '" );
				sbError.Append( subjDetail.ClinicalTrialName );
				sbError.Append( "' does not have site code '" );
				sbError.Append( subjDetail.Site );
				sbError.Append( "' attached to it.");
				log.Error( sbError.ToString(), ex );
				return false;
			}

			// check study / site / subject label
			try
			{
				subjDetail.SubjectId = bufferDb.GetSubjectId( subjDetail.ClinicalTrialId, subjDetail.Site, subjDetail.SubjectLabel );
			}
			catch(Exception ex)
			{
				// study /site / subject do not exist
				subjDetail.SubjectResponseStatus = BufferAPI.BufferResponseStatus.FailSubjectLabel;
				// log error
				StringBuilder sbError = new StringBuilder();
				sbError.Append( "Study '" );
				sbError.Append( subjDetail.ClinicalTrialName );
				sbError.Append( "' / site '" );
				sbError.Append( subjDetail.Site );
				sbError.Append( "' does not contain subject '");
				sbError.Append( subjDetail.SubjectLabel );
				sbError.Append( "'." );
				log.Error( sbError.ToString(), ex );
				return false;
			}

			// going to collect study details for use with each response (+ visit /eform)
			StudyInfo studyInfo = new StudyInfo();
			try
			{
				studyInfo = bufferDb.GetStudyInfo( subjDetail.ClinicalTrialId );
			}
			catch(Exception ex)
			{
				// study /site / subject do not exist
				subjDetail.SubjectResponseStatus = BufferAPI.BufferResponseStatus.FailDbConnection;
				// log error
				StringBuilder sbError = new StringBuilder();
				sbError.Append( "Study '" );
				sbError.Append( subjDetail.ClinicalTrialName );
				sbError.Append( "' / site '" );
				sbError.Append( subjDetail.Site );
				sbError.Append( "' does not contain subject '");
				sbError.Append( subjDetail.SubjectLabel );
				sbError.Append( "'." );
				log.Error( sbError.ToString(), ex );
				return false;
			}


			// for each response
			for(int i=0; i<subjDetail.Responses.Count; i++)
			{
				// get study / data item code
				((ResponseDetails)subjDetail.Responses[i]).DataItemId = studyInfo.GetDataItemId( ((ResponseDetails)subjDetail.Responses[i]).DataItemCode );
				// check study / data item code
				if(((ResponseDetails)subjDetail.Responses[i]).DataItemId == BufferAPI._DEFAULT_MISSING_NUMERIC)
				{
					// study / data item doesn't exist
					subjDetail.SubjectResponseStatus = BufferAPI.BufferResponseStatus.FailDataItemCode;
					((ResponseDetails)subjDetail.Responses[i]).BufferResponseStatus = BufferAPI.BufferResponseStatus.FailDataItemCode;
					// log error
					StringBuilder sbError = new StringBuilder();
					sbError.Append( "Study  '" );
					sbError.Append( subjDetail.ClinicalTrialName );
					sbError.Append( "' does not contain data item '" );
					sbError.Append( ((ResponseDetails)subjDetail.Responses[i]).DataItemCode );
					sbError.Append ( "'." );
					log.Error( sbError.ToString() );
					return false;
				}

				// check response value is of correct data type
				try
				{
					CheckDataItemResponseValue( studyInfo.GetResponseDataItem( ((ResponseDetails)subjDetail.Responses[i]).DataItemId ), ((ResponseDetails)subjDetail.Responses[i]).ResponseValue);
				}
				catch(Exception ex)
				{
					// response data type is invalid
					if(subjDetail.SubjectResponseStatus == BufferAPI.BufferResponseStatus.Success)
					{
						subjDetail.SubjectResponseStatus = BufferAPI.BufferResponseStatus.FailResponseValue;
					}
					((ResponseDetails)subjDetail.Responses[i]).BufferResponseStatus = BufferAPI.BufferResponseStatus.FailResponseValue;
					// log
					log.Error(ex.Message);
					bOk = false;
				}

				// VisitId
				((ResponseDetails)subjDetail.Responses[i]).VisitId = studyInfo.GetVisitId( ((ResponseDetails)subjDetail.Responses[i]).VisitCode );
				// CrfPageId
				((ResponseDetails)subjDetail.Responses[i]).CRFPageId = studyInfo.GetEformId( ((ResponseDetails)subjDetail.Responses[i]).CRFPageCode );
			}
			return bOk;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="responseDataItem"></param>
		/// <param name="sDataItemValue"></param>
		private static void CheckDataItemResponseValue(ResponseDataItem responseDataItem, string sDataItemValue)
		{
			// now check data value according to type
			// check has no illegal chars `"|~
			string sIllegal = "`|~\"";
			char[] chIllegal = sIllegal.ToCharArray();
			if(sDataItemValue.IndexOfAny(chIllegal) > -1)
			{
				throw ( new Exception( "Response contains an illegal character." ) );
			}

			switch(responseDataItem.MACRODataType)
			{
				case BufferAPI.MACRODataTypes.Text:
				{
					// check is not too long
					if(sDataItemValue.Length > responseDataItem.Length)
					{
						throw ( new Exception( "Response is longer than " + responseDataItem.Length.ToString() + " characters." ) );
					}
					// check string format
					if(responseDataItem.Format != "")
					{
						// convert dataitemvalue to 9/A MACRO format
						StringBuilder sbValue = new StringBuilder();
						for(int i=0; i < sDataItemValue.Length; i++)
						{
							if(Char.IsLetter(sDataItemValue[i]))
							{
								sbValue.Append("A");
							}
							else if(Char.IsDigit(sDataItemValue[i]))
							{
								sbValue.Append("9");
							}
							else
							{
								sbValue.Append(sDataItemValue[i]);
							}
						}
						// check macro format is reflected in dataitem
						if( responseDataItem.Format.ToUpper() != sbValue.ToString() )
						{
							throw ( new Exception( "Response does not match specified format." ) );
						}
					}
					break;
				}
				case BufferAPI.MACRODataTypes.IntegerData:
				{
					// check numeric & integer
					try
					{
						int nTestInteger = Convert.ToInt32(sDataItemValue);
					}
					catch
					{
						throw ( new Exception ( "Response is not an integer value.") );
					}
					// check integer format
					try
					{
						CheckNumber(sDataItemValue, responseDataItem.Format);
					}
					catch
					{
						throw ( new Exception ( "Response is invalid for the data item format.") );
					}
					break;
				}
				case BufferAPI.MACRODataTypes.Real:
				case BufferAPI.MACRODataTypes.LabTest:
				{
					// check real data type formatting
					CheckNumber(sDataItemValue, responseDataItem.Format);
					break;
				}
				case BufferAPI.MACRODataTypes.Category:
				{
					// check category exists in list
					bool bCatExists = false;
					// check both codes & values					
					foreach(CategoryItem catItem in responseDataItem.Categories)
					{
						if((sDataItemValue == catItem.Code)||(sDataItemValue == catItem.Value))
						{
							bCatExists = true;
							break;
						}
					}
					// if category not exist throw exception
					if(!bCatExists)
					{
						throw ( new Exception ( "Category response not in category code/value list.") );
					}
					break;
				}
				case BufferAPI.MACRODataTypes.Date:
				{
					// check date with expected input format
					// parse inputted date/time
					// get expected format
					string sFormat = BufferBasicChecks.DateFormatString(sDataItemValue);
					// check date/time
					try
					{
						// DPH 22/11/2007 - set culture as local culture is disrupting output of DateTime.ParseExact
						IFormatProvider culture = new System.Globalization.CultureInfo("en-GB", true);
						// pass in GB culture (as expect dd/MM/yyyy format)
						DateTime dtDate = DateTime.ParseExact(sDataItemValue,sFormat,culture);
					}
					catch
					{
						throw ( new Exception( "Response date/time is an invalid accepted format. It must be of format dd/mm/yyyy or similar." ) );
					}
					break;
				}
				case BufferAPI.MACRODataTypes.Multimedia:
				{
					throw ( new Exception( "Response is a multimedia data type. This cannot be handled by the buffer API." ) );
				}
				case BufferAPI.MACRODataTypes.Thesaurus:
				{
					throw ( new Exception( "Response is a thesaurus data type. This cannot be handled by the buffer API." ) );
				}
			}
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="sCheckNumber"></param>
		/// <param name="sFormat"></param>
		private static void CheckNumber(string sCheckNumber, string sFormat)
		{
			// to check number against format
			// convert format and number to doubles
			double dblCheckNumber = Convert.ToDouble(sCheckNumber);
			// if format is passed
			if(sFormat != "")
			{
				// replace any # with 9 to give a format maximum
				double dblFormat = Convert.ToDouble(sFormat.Replace("#","9"));

				// compare numbers
				if(dblFormat > 0)
				{
					if((dblCheckNumber < 0)||(dblCheckNumber > dblFormat))
					{
						throw ( new Exception( "Response is invalid for the data item format." ) );
					}
				}
				else
				{
					// negative
					if((dblCheckNumber > 0)||(dblCheckNumber < dblFormat))
					{
						throw ( new Exception( "Response is invalid for the data item format." ) );
					}
				}

				// decimal places?
				// format
				string sDecSeparator = ".";
				int nFormatDecPlace = sFormat.IndexOf(sDecSeparator);
				int nFormatDecPlaces = 0;
				if( nFormatDecPlace > -1 ) 
				{
					nFormatDecPlaces = sFormat.Length - ( nFormatDecPlace + sDecSeparator.Length );
				}
					
				if(nFormatDecPlaces > 0)
				{
					// data
					int nDataDecPlace = sCheckNumber.IndexOf(sDecSeparator);
					int nDataDecPlaces = 0;
					if( nDataDecPlace > -1 ) 
					{
						nDataDecPlaces = sCheckNumber.Length - ( nDataDecPlace + sDecSeparator.Length );
					}
					if(nDataDecPlaces > nFormatDecPlaces)
					{
						throw ( new Exception( "Response is invalid for the data item format." ) );
					}
				}
			}
		}

		/// <summary>
		/// Take a date string passed to api
		/// One of formats:-
		/// dd/mm/yyyy
		/// dd/mm/yyyy hh:mm
		/// dd/mm/yyyy hh:mm:ss
		/// hh:mm
		/// hh:mm:ss
		/// return expected format string
		/// </summary>
		/// <param name="sDate"></param>
		/// <returns></returns>
		public static string DateFormatString(string sDate)
		{
			string sFormat = "";
			if(sDate.IndexOf("/") > -1)
			{
				sFormat = "dd/MM/yyyy";
			}	
			string sTimeSep = ":";
			// check if need to include time
			switch(sDate.Split(sTimeSep.ToCharArray()).Length)
			{
				case 2:
				{
					//hh:mm
					sFormat += (sFormat.Length==0)?"":" ";
					sFormat = sFormat + "HH:mm";
					break;
				}
				case 3:
				{
					//hh:mm:ss
					sFormat += (sFormat.Length==0)?"":" ";
					sFormat = sFormat + "HH:mm:ss";
					break;
				}
			}
			return sFormat;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="sMACRODateFormat"></param>
		/// <returns></returns>
		public static string FormatMACROToMicrosoftDate(string sMACRODateFormat)
		{
			// possible MACRO date formats
			//  Dates:
			//
			//      YMD
			//      DMY
			//      MDY
			//      YM
			//      MY
			//
			//  Times:
			//
			//      HMS
			//      HM
			//
			//  Combinations:
			//      YMD HMS
			//      DMY HMS
			//      MDY HMS
			//      YMD HM
			//      DMY HM
			//      MDY HM
			
			string sFormat = sMACRODateFormat;
			string sNewFormat = "";

			// make macro date an above format
			// Get the basic format from the mask. see above list for allowable values.
			sFormat=sFormat.Replace("d","D");
			sFormat=sFormat.Replace("DD","D");
			sFormat=sFormat.Replace("m","M");
			sFormat=sFormat.Replace("MM","M");
			sFormat=sFormat.Replace("y","Y");
			sFormat=sFormat.Replace("YYYY","Y");
			sFormat=sFormat.Replace("YY","Y");
			sFormat=sFormat.Replace("h","H");
			sFormat=sFormat.Replace("HH","H");
			sFormat=sFormat.Replace("s","S");
			sFormat=sFormat.Replace("SS","S");
			sFormat=sFormat.Replace(" ","");
			// Get the format into a standard string, based on any of the following: "/.:-" and whitespace.
			// "/.:-" are the possible separators
			sFormat=sFormat.Replace(".","");
			sFormat=sFormat.Replace(":","");
			sFormat=sFormat.Replace("-","");
			sFormat=sFormat.Replace("/","");

			switch(sFormat)
			{
				case "YMD":
				{
					int nSep1 = ReturnNextSeparatorPos(sMACRODateFormat, 0);
					int nSep2 = ReturnNextSeparatorPos(sMACRODateFormat, nSep1 + 1);
					sNewFormat = ((nSep1)==2)?"yy":"yyyy";
					sNewFormat += sMACRODateFormat.Substring(nSep1,1);
					sNewFormat += (DPLength(nSep1,nSep2)==1)?"M":"MM";
					sNewFormat += sMACRODateFormat.Substring(nSep2,1);
					sNewFormat += (DPLength(nSep2,sMACRODateFormat.Length)==1)?"d":"dd";
					break;
				}
				case "DMY":
				{
					int nSep1 = ReturnNextSeparatorPos(sMACRODateFormat, 0);
					int nSep2 = ReturnNextSeparatorPos(sMACRODateFormat, nSep1 + 1);
					sNewFormat = ((nSep1)==1)?"d":"dd";
					sNewFormat += sMACRODateFormat.Substring(nSep1,1);
					sNewFormat += (DPLength(nSep1,nSep2)==1)?"M":"MM";
					sNewFormat += sMACRODateFormat.Substring(nSep2,1);
					sNewFormat += (DPLength(nSep2,sMACRODateFormat.Length)==2)?"yy":"yyyy";
					break;
				}
				case "MDY":
				{
					int nSep1 = ReturnNextSeparatorPos(sMACRODateFormat, 0);
					int nSep2 = ReturnNextSeparatorPos(sMACRODateFormat, nSep1 + 1);
					sNewFormat = ((nSep1)==1)?"M":"MM";
					sNewFormat += sMACRODateFormat.Substring(nSep1,1);
					sNewFormat += (DPLength(nSep1,nSep2)==1)?"d":"dd";
					sNewFormat += sMACRODateFormat.Substring(nSep2,1);
					sNewFormat += (DPLength(nSep2,sMACRODateFormat.Length)==2)?"yy":"yyyy";
					break;
				}
				case "YM":
				{
					int nSep1 = ReturnNextSeparatorPos(sMACRODateFormat, 0);
					sNewFormat = ((nSep1)==2)?"yy":"yyyy";
					sNewFormat += sMACRODateFormat.Substring(nSep1,1);
					sNewFormat += (DPLength(nSep1,sMACRODateFormat.Length)==1)?"M":"MM";
					break;
				}
				case "MY":
				{
					int nSep1 = ReturnNextSeparatorPos(sMACRODateFormat, 0);
					sNewFormat = ((nSep1)==1)?"M":"MM";
					sNewFormat += sMACRODateFormat.Substring(nSep1,1);
					sNewFormat += (DPLength(nSep1,sMACRODateFormat.Length)==2)?"yy":"yyyy";
					break;
				}
				case "HMS":
				{
					int nSep1 = ReturnNextSeparatorPos(sMACRODateFormat, 0);
					int nSep2 = ReturnNextSeparatorPos(sMACRODateFormat, nSep1 + 1);
					sNewFormat = ((nSep1)==1)?"H":"HH";
					sNewFormat += sMACRODateFormat.Substring(nSep1,1);
					sNewFormat += (DPLength(nSep1,nSep2)==1)?"m":"mm";
					sNewFormat += sMACRODateFormat.Substring(nSep2,1);
					sNewFormat += (DPLength(nSep2,sMACRODateFormat.Length)==1)?"s":"ss";
					break;
				}
				case "HM":
				{
					int nSep1 = ReturnNextSeparatorPos(sMACRODateFormat, 0);
					sNewFormat = ((nSep1)==1)?"H":"HH";
					sNewFormat += sMACRODateFormat.Substring(nSep1,1);
					sNewFormat += (DPLength(nSep1,sMACRODateFormat.Length)==1)?"m":"mm";
					break;
				}
				case "YMDHMS":
				{
					int nSep1 = ReturnNextSeparatorPos(sMACRODateFormat, 0);
					int nSep2 = ReturnNextSeparatorPos(sMACRODateFormat, nSep1 + 1);
					int nSep3 = ReturnNextSeparatorPos(sMACRODateFormat, nSep2 + 1);
					int nSep4 = ReturnNextSeparatorPos(sMACRODateFormat, nSep3 + 1);
					int nSep5 = ReturnNextSeparatorPos(sMACRODateFormat, nSep4 + 1);
					sNewFormat = ((nSep1)==2)?"yy":"yyyy";
					sNewFormat += sMACRODateFormat.Substring(nSep1,1);
					sNewFormat += (DPLength(nSep1,nSep2)==1)?"M":"MM";
					sNewFormat += sMACRODateFormat.Substring(nSep2,1);
					sNewFormat += (DPLength(nSep2,nSep3)==1)?"d":"dd";
					sNewFormat += sMACRODateFormat.Substring(nSep3,1);
					sNewFormat += (DPLength(nSep3,nSep4)==1)?"H":"HH";
					sNewFormat += sMACRODateFormat.Substring(nSep4,1);
					sNewFormat += (DPLength(nSep4,nSep5)==1)?"m":"mm";
					sNewFormat += sMACRODateFormat.Substring(nSep5,1);
					sNewFormat += (DPLength(nSep5,sMACRODateFormat.Length)==1)?"s":"ss";
					break;
				}
				case "DMYHMS":
				{
					int nSep1 = ReturnNextSeparatorPos(sMACRODateFormat, 0);
					int nSep2 = ReturnNextSeparatorPos(sMACRODateFormat, nSep1 + 1);
					int nSep3 = ReturnNextSeparatorPos(sMACRODateFormat, nSep2 + 1);
					int nSep4 = ReturnNextSeparatorPos(sMACRODateFormat, nSep3 + 1);
					int nSep5 = ReturnNextSeparatorPos(sMACRODateFormat, nSep4 + 1);
					sNewFormat = ((nSep1)==1)?"d":"dd";
					sNewFormat += sMACRODateFormat.Substring(nSep1,1);
					sNewFormat += (DPLength(nSep1,nSep2)==1)?"M":"MM";
					sNewFormat += sMACRODateFormat.Substring(nSep2,1);
					sNewFormat += (DPLength(nSep2,nSep3)==2)?"yy":"yyyy";
					sNewFormat += sMACRODateFormat.Substring(nSep3,1);
					sNewFormat += (DPLength(nSep3,nSep4)==1)?"H":"HH";
					sNewFormat += sMACRODateFormat.Substring(nSep4,1);
					sNewFormat += (DPLength(nSep4,nSep5)==1)?"m":"mm";
					sNewFormat += sMACRODateFormat.Substring(nSep5,1);
					sNewFormat += (DPLength(nSep5,sMACRODateFormat.Length)==1)?"s":"ss";
					break;
				}
				case "MDYHMS":
				{
					int nSep1 = ReturnNextSeparatorPos(sMACRODateFormat, 0);
					int nSep2 = ReturnNextSeparatorPos(sMACRODateFormat, nSep1 + 1);
					int nSep3 = ReturnNextSeparatorPos(sMACRODateFormat, nSep2 + 1);
					int nSep4 = ReturnNextSeparatorPos(sMACRODateFormat, nSep3 + 1);
					int nSep5 = ReturnNextSeparatorPos(sMACRODateFormat, nSep4 + 1);
					sNewFormat = ((nSep1)==1)?"M":"MM";
					sNewFormat += sMACRODateFormat.Substring(nSep1,1);
					sNewFormat += (DPLength(nSep1,nSep2)==1)?"d":"dd";
					sNewFormat += sMACRODateFormat.Substring(nSep2,1);
					sNewFormat += (DPLength(nSep2,nSep3)==2)?"yy":"yyyy";
					sNewFormat += sMACRODateFormat.Substring(nSep3,1);
					sNewFormat += (DPLength(nSep3,nSep4)==1)?"H":"HH";
					sNewFormat += sMACRODateFormat.Substring(nSep4,1);
					sNewFormat += (DPLength(nSep4,nSep5)==1)?"m":"mm";
					sNewFormat += sMACRODateFormat.Substring(nSep5,1);
					sNewFormat += (DPLength(nSep5,sMACRODateFormat.Length)==1)?"s":"ss";
					break;
				}
				case "YMDHM":
				{
					int nSep1 = ReturnNextSeparatorPos(sMACRODateFormat, 0);
					int nSep2 = ReturnNextSeparatorPos(sMACRODateFormat, nSep1 + 1);
					int nSep3 = ReturnNextSeparatorPos(sMACRODateFormat, nSep2 + 1);
					int nSep4 = ReturnNextSeparatorPos(sMACRODateFormat, nSep3 + 1);
					sNewFormat = ((nSep1)==2)?"yy":"yyyy";
					sNewFormat += sMACRODateFormat.Substring(nSep1,1);
					sNewFormat += (DPLength(nSep1,nSep2)==1)?"M":"MM";
					sNewFormat += sMACRODateFormat.Substring(nSep2,1);
					sNewFormat += (DPLength(nSep2,nSep3)==1)?"d":"dd";
					sNewFormat += sMACRODateFormat.Substring(nSep3,1);
					sNewFormat += (DPLength(nSep3,nSep4)==1)?"H":"HH";
					sNewFormat += sMACRODateFormat.Substring(nSep4,1);
					sNewFormat += (DPLength(nSep4,sMACRODateFormat.Length)==1)?"m":"mm";
					break;
				}
				case "DMYHM":
				{
					int nSep1 = ReturnNextSeparatorPos(sMACRODateFormat, 0);
					int nSep2 = ReturnNextSeparatorPos(sMACRODateFormat, nSep1 + 1);
					int nSep3 = ReturnNextSeparatorPos(sMACRODateFormat, nSep2 + 1);
					int nSep4 = ReturnNextSeparatorPos(sMACRODateFormat, nSep3 + 1);
					sNewFormat = ((nSep1)==1)?"d":"dd";
					sNewFormat += sMACRODateFormat.Substring(nSep1,1);
					sNewFormat += (DPLength(nSep1,nSep2)==1)?"M":"MM";
					sNewFormat += sMACRODateFormat.Substring(nSep2,1);
					sNewFormat += (DPLength(nSep2,nSep3)==2)?"yy":"yyyy";
					sNewFormat += sMACRODateFormat.Substring(nSep3,1);
					sNewFormat += (DPLength(nSep3,nSep4)==1)?"H":"HH";
					sNewFormat += sMACRODateFormat.Substring(nSep4,1);
					sNewFormat += (DPLength(nSep4,sMACRODateFormat.Length)==1)?"m":"mm";
					break;
				}
				case "MDYHM":
				{
					int nSep1 = ReturnNextSeparatorPos(sMACRODateFormat, 0);
					int nSep2 = ReturnNextSeparatorPos(sMACRODateFormat, nSep1 + 1);
					int nSep3 = ReturnNextSeparatorPos(sMACRODateFormat, nSep2 + 1);
					int nSep4 = ReturnNextSeparatorPos(sMACRODateFormat, nSep3 + 1);
					sNewFormat = ((nSep1)==1)?"M":"MM";
					sNewFormat += sMACRODateFormat.Substring(nSep1,1);
					sNewFormat += (DPLength(nSep1,nSep2)==1)?"d":"dd";
					sNewFormat += sMACRODateFormat.Substring(nSep2,1);
					sNewFormat += (DPLength(nSep2,nSep3)==2)?"yy":"yyyy";
					sNewFormat += sMACRODateFormat.Substring(nSep3,1);
					sNewFormat += (DPLength(nSep3,nSep4)==1)?"H":"HH";
					sNewFormat += sMACRODateFormat.Substring(nSep4,1);
					sNewFormat += (DPLength(nSep4,sMACRODateFormat.Length)==1)?"m":"mm";
					break;
				}
			}
			
			return sNewFormat;
		}

		private static int ReturnNextSeparatorPos(string sDateFormat, int nLastPos)
		{
			string sSeps = "/.:- ";
			char[] aSeps = sSeps.ToCharArray();
			int nNextSep = sDateFormat.IndexOfAny(aSeps, nLastPos);
			return nNextSep;
		}

		private static int DPLength(int nSepA, int nSepB)
		{
			return( nSepB - nSepA - 1);
		}
	}
}
