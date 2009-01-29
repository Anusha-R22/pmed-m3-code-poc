using System;
using System.IO;
using System.Text;
using System.Xml;
using log4net;

namespace InferMed.MACROBuffer
{
	/// <summary>
	/// Summary description for BufferMessageXML.
	/// </summary>
	class BufferMessageXML
	{
		// member variables
		private XmlDocument _messageXML; 
		// log4net
		private static readonly ILog log = LogManager.GetLogger( typeof(MACROBufferAPI) );

		/// <summary>
		/// 
		/// </summary>
		public BufferMessageXML()
		{
			_messageXML = new XmlDocument(); 
		}

		// parse the message ensuring compliant with what is expected
		/// <summary>
		/// 
		/// </summary>
		/// <param name="sBufferMessage"></param>
		/// <param name="subjDetail"></param>
		/// <returns></returns>
		public BufferAPI.BufferResponseStatus ParseBufferMessage(string sBufferMessage, ref SubjectDetails subjDetail)
		{
			BufferAPI.BufferResponseStatus eMessageStatus = BufferAPI.BufferResponseStatus.Success;
			string sAttributeData = "";

			try
			{
				// load xml string
				_messageXML.LoadXml(sBufferMessage);

				// check for multiples of nodes which should be single
				// only 1 ODM
				int nCountODM = 0;
				nCountODM = _messageXML.SelectNodes( "//ODM" ).Count;
				if(nCountODM != 1)
				{
					throw ( new Exception( "There should be exactly 1 ODM node." ) );
				}
				// only 1 ClinicalData
				int nCountClinicalData = 0;
				nCountClinicalData = _messageXML.SelectNodes( "//ODM/ClinicalData" ).Count;
				if(nCountClinicalData != 1)
				{
					throw ( new Exception( "There should be exactly 1 ClinicalData node." ) );
				}
				// only 1 SubjectData
				int nCountSubjectData = 0;
				nCountSubjectData = _messageXML.SelectNodes( "//ODM/ClinicalData/SubjectData" ).Count;
				if(nCountSubjectData != 1)
				{
					throw ( new Exception( "There should be exactly 1 SubjectData node." ) );
				}

				// parse ODM Node (1x)
				XmlNode nodeODM = _messageXML.SelectSingleNode( "//ODM" );
				// check attributes
				// FileType - do nothing with
				sAttributeData = GetXmlAttribute(nodeODM, "FileType");
				// FileOID - do nothing with
				sAttributeData = GetXmlAttribute(nodeODM, "FileOID");
				// CreationDateTime - do nothing with
				sAttributeData = GetXmlAttribute(nodeODM, "CreationDateTime");
				// free memory
				nodeODM = null;

				// parse ClinicalData Node(1x)
				XmlNode nodeClinicalData = _messageXML.SelectSingleNode( "//ODM/ClinicalData" );
				// check attributes
				// StudyOID - Study code
				sAttributeData = GetXmlAttribute(nodeClinicalData, "StudyOID");
				// store in SubjectDetails
				subjDetail.ClinicalTrialName = sAttributeData;
				// MetaDataVersionOID - do nothing with
				sAttributeData = GetXmlAttribute(nodeClinicalData, "MetaDataVersionOID");
				// free memory
				nodeClinicalData = null;

				// parse SubjectData (1x)
				// get SubjectData node
				XmlNode nodeSubjectData = _messageXML.SelectSingleNode( "//ODM/ClinicalData/SubjectData" );
				// SubjectKey - Subject Label
				sAttributeData = GetXmlAttribute(nodeSubjectData, "SubjectKey");
				// store in SubjectDetails
				subjDetail.SubjectLabel = sAttributeData;

				// only 1 SiteRef
				int nCountSiteRef = 0;
				nCountSiteRef = nodeSubjectData.SelectNodes( "//SubjectData/SiteRef" ).Count;
				if(nCountSiteRef != 1)
				{
					throw ( new Exception( "There should be exactly 1 SiteRef node." ) );
				}

				// parse SiteRef (1x)
				XmlNode nodeSiteRef = nodeSubjectData.SelectSingleNode( "//SubjectData/SiteRef" );
				// LocationOID - Site
				sAttributeData = GetXmlAttribute(nodeSiteRef, "LocationOID");
				// store in SubjectDetails
				subjDetail.Site = sAttributeData;

				// at least 1 StudyEventData
				int nCountStudyEventData = 0;
				nCountStudyEventData = nodeSubjectData.SelectNodes( "//SubjectData/StudyEventData" ).Count;
				if(nCountStudyEventData < 1)
				{
					throw ( new Exception( "There should be at least 1 StudyEventData node." ) );
				}

				// loop through visits
				foreach ( XmlNode nodeStudyEventData in nodeSubjectData.SelectNodes( "//SubjectData/StudyEventData" ) )
				{
					// parse visits
					ParseVisits( nodeStudyEventData, ref subjDetail);
				}
			}
			catch(Exception ex)
			{
				string sErrorMessage = ex.ToString();
				// an error has occurred
				eMessageStatus = BufferAPI.BufferResponseStatus.FailParse;
				// store in log file
				log.Error("Parse message failure", ex);
			}
			
			return eMessageStatus;
		}

		/// <summary>
		/// Parse the visits xml
		/// </summary>
		/// <param name="nodeStudyEventData"></param>
		/// <param name="subjDetail"></param>
		private void ParseVisits(XmlNode nodeStudyEventData, ref SubjectDetails subjDetail)
		{
			// parse StudyEventData (X)
			// StudyEventOID - Visit Code
			string sVisitCode = GetXmlAttribute(nodeStudyEventData, "StudyEventOID");
			// StudyEventRepeatKey - Visit Repeat Number
			string sAttributeData = GetXmlAttribute(nodeStudyEventData, "StudyEventRepeatKey", true);
			int nVisitCycle = BufferAPI._DEFAULT_MISSING_NUMERIC;
			if(sAttributeData != "")
			{
				nVisitCycle = Convert.ToInt16( sAttributeData );
			}

			// at least 1 FormData
			int nCountFormData = 0;
			nCountFormData = nodeStudyEventData.SelectNodes( "FormData" ).Count;
			if(nCountFormData < 1)
			{
				throw ( new Exception( "There should be at least 1 FormData node." ) );
			}

			// loop through eForms
			foreach ( XmlNode nodeFormData in nodeStudyEventData.SelectNodes( "FormData" ) )
			{
				// parse FormData (X)
				// FormOID - eForm Code
				string sEFormCode = GetXmlAttribute(nodeFormData, "FormOID");
				// FormRepeatKey - eForm Repeat Number
				sAttributeData = GetXmlAttribute(nodeFormData, "FormRepeatKey", true);
				int nEFormCycle = BufferAPI._DEFAULT_MISSING_NUMERIC;
				if(sAttributeData != "")
				{
					nEFormCycle = Convert.ToInt16( sAttributeData );
				}

				// at least 1 ItemGroupData
				int nCountItemGroupData = 0;
				nCountItemGroupData = nodeFormData.SelectNodes( "ItemGroupData" ).Count;
				if(nCountItemGroupData < 1)
				{
					throw ( new Exception( "There should be at least 1 ItemGroupData node." ) );
				}

				foreach( XmlNode nodeItemGroupData in nodeFormData.SelectNodes( "ItemGroupData" ) )
				{
					// parse ItemGroupData (X)
					// ItemGroupOID - Item Group Code
					string sItemGroupCode = GetXmlAttribute(nodeItemGroupData, "ItemGroupOID");
					// ItemGroupRepeatKey - Repeat Number
					sAttributeData = GetXmlAttribute(nodeItemGroupData, "ItemGroupRepeatKey", true);
					int nItemGroupRepeat = BufferAPI._DEFAULT_MISSING_NUMERIC;
					if(sAttributeData != "")
					{
						nItemGroupRepeat = Convert.ToInt16( sAttributeData );
					}

					// at least 1 ItemData
					int nCountItemData = 0;
					nCountItemData = nodeItemGroupData.SelectNodes( "ItemData" ).Count;
					if(nCountItemData < 1)
					{
						throw ( new Exception( "There should be at least 1 ItemData node." ) );
					}
					// loop through data items
					foreach( XmlNode nodeItemData in nodeItemGroupData.SelectNodes( "ItemData" ) )
					{
						// parse ItemData (X)
						// ItemOID - Data Item Code
						string sDataItemCode = GetXmlAttribute(nodeItemData, "ItemOID");
						// Value - Data Item Response Value
						string sDataItemValue = GetXmlAttribute(nodeItemData, "Value");

						// check only 0 or 1 audit record
						int nCountAuditRecord = 0;
						nCountAuditRecord = nodeItemData.SelectNodes( "AuditRecord" ).Count;
						if(nCountAuditRecord > 1)
						{
							throw ( new Exception( "There should be 0 or 1 AuditRecord nodes." ) );
						}

						DateTime dtData = DateTime.FromOADate(0);

						if( nCountAuditRecord == 1)
						{

							XmlNode nodeAuditRecord = nodeItemData.SelectSingleNode( "AuditRecord" );
						
							// only 1 UserRef
							int nCountUserRef = 0;
							nCountUserRef = nodeAuditRecord.SelectNodes( "UserRef" ).Count;
							if(nCountUserRef != 1)
							{
								throw ( new Exception( "There should be exactly 1 UserRef node." ) );
							}
							XmlNode nodeUserRef = nodeAuditRecord.SelectSingleNode( "UserRef" );
							// UserOID - Not Used in MACRO
							string sUserOID = GetXmlAttribute(nodeUserRef, "UserOID");

							// only 1 LocationRef
							int nCountLocationRef = 0;
							nCountLocationRef = nodeAuditRecord.SelectNodes( "LocationRef" ).Count;
							if(nCountLocationRef != 1)
							{
								throw ( new Exception( "There should be exactly 1 LocationRef node." ) );
							}
							XmlNode nodeLocationRef = nodeAuditRecord.SelectSingleNode( "LocationRef" );
							// LocationOID - Not Used in MACRO
							string sLocationOID = GetXmlAttribute(nodeLocationRef, "LocationOID");

							// only 1 DateTimeStamp
							int nCountDateTimeStamp = 0;
							nCountDateTimeStamp = nodeAuditRecord.SelectNodes( "DateTimeStamp" ).Count;
							if(nCountDateTimeStamp != 1)
							{
								throw ( new Exception( "There should be exactly 1 DateTimeStamp node." ) );
							}

							XmlNode nodeDataTimeStamp = nodeAuditRecord.SelectSingleNode( "DateTimeStamp" );

							// parse DateTimeStamp (1x)
							string sDataTimeStamp = "";
							if(nodeDataTimeStamp.HasChildNodes)
							{
								if(nodeDataTimeStamp.FirstChild.NodeType == XmlNodeType.Text)
								{
									sDataTimeStamp = (nodeDataTimeStamp.FirstChild.Value!=null)?nodeDataTimeStamp.FirstChild.Value:"";
								}
							}
							if(sDataTimeStamp != "")
							{
								// parse inputted date/time
								string sFormat = BufferBasicChecks.DateFormatString(sDataTimeStamp);
								// check if need to include time
								try
								{
									// DPH 22/11/2007 - set culture as local culture is disrupting output of DateTime.ParseExact
									IFormatProvider culture = new System.Globalization.CultureInfo("en-GB", true);
									// pass in GB culture (as expect dd/MM/yyyy format)
									dtData = DateTime.ParseExact(sDataTimeStamp,sFormat,culture);
								}
								catch
								{
									// store in log file
									log.Warn("DateTime Parse exact error: DateTimeStamp - " + sDataTimeStamp + " Format - " + sFormat);
									throw ( new Exception( "The given data timestamp was of an invalid format." ) );
								}
							}
						}
						// create ReponseDetails & add to Subject details
						ResponseDetails responseDetail = new ResponseDetails();
						responseDetail.VisitCode = sVisitCode;
						responseDetail.VisitCycle = nVisitCycle;
						responseDetail.CRFPageCode = sEFormCode;
						responseDetail.CRFPageCycle = nEFormCycle;
						responseDetail.DataItemCode = sDataItemCode;
						if( nItemGroupRepeat != BufferAPI._DEFAULT_MISSING_NUMERIC )
						{
							if ( nItemGroupRepeat == 0 )
							{
								nItemGroupRepeat++;
							}
							responseDetail.DataItemRepeatNo = nItemGroupRepeat;
						}
						responseDetail.ResponseValue = sDataItemValue;
						responseDetail.DataCompareDate = dtData;

						// add response row to SubjectDetails
						subjDetail.AddResponse( responseDetail );
					}
				}
			}

		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="oXmlNode"></param>
		/// <param name="sAttributeName"></param>
		/// <returns></returns>
		private string GetXmlAttributeNoError(XmlNode oXmlNode, string sAttributeName)
		{
			// get attribute
			string sValue = "";
			if(oXmlNode.Attributes[sAttributeName] != null)
			{
				sValue = oXmlNode.Attributes[sAttributeName].Value.ToString();
			}
			return sValue;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="oXmlNode"></param>
		/// <param name="sAttributeName"></param>
		/// <returns></returns>
		private string GetXmlAttribute(XmlNode oXmlNode, string sAttributeName)
		{
			// get attribute
			return oXmlNode.Attributes[sAttributeName].Value.ToString();
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="oXmlNode"></param>
		/// <param name="sAttributeName"></param>
		/// <param name="bOptionalAttribute"></param>
		/// <returns></returns>
		private string GetXmlAttribute(XmlNode oXmlNode, string sAttributeName, bool bOptionalAttribute)
		{
			string sValue = "";
			try
			{
				sValue = GetXmlAttribute(oXmlNode, sAttributeName);
			}
			catch(Exception ex)
			{
				// if not optional attribute throw error
				if(!bOptionalAttribute)
				{
					throw(new Exception(ex.Message, ex.InnerException));
				}
			}
			return sValue;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="subjDetail"></param>
		/// <returns></returns>
		public string CreateMessageXMLReport(ref SubjectDetails subjDetail, bool bMACROCommit)
		{
			string sReportXml = "";
			try
			{
				// create xml in memory
				MemoryStream oMemStream = new MemoryStream();
				// create xmltextwriter - ASCII encoding
				//XmlTextWriter oXmlTw = new XmlTextWriter(oMemStream, Encoding.ASCII);
				XmlTextWriter oXmlTw = new XmlTextWriter(oMemStream, Encoding.GetEncoding("ISO-8859-1"));
				
				// Create Message report
				// start document
				oXmlTw.WriteStartDocument();
				// comment
				oXmlTw.WriteComment("Generated by MACRO Buffer API Application "+DateTime.Now);
				// report main element
				oXmlTw.WriteStartElement("BufferMessageReport");
				// attributes
				// Buffer Save Status
				oXmlTw.WriteAttributeString("BufferSaveStatus", ((int)subjDetail.SubjectResponseStatus).ToString());

				// Buffer Message Errors element
				oXmlTw.WriteStartElement("MessageErrors");

				if( subjDetail.SubjectResponseStatus > 0 )
				{
					// for all message error types
					//  open message error
					oXmlTw.WriteStartElement("MessageError");
					oXmlTw.WriteAttributeString("BufferErrorStatus", ((int)subjDetail.SubjectResponseStatus).ToString());
					string sErrorDesc = ErrorDescription(subjDetail.SubjectResponseStatus);

					oXmlTw.WriteAttributeString("BufferErrorDescription", sErrorDesc);
					oXmlTw.WriteEndElement();
					// if response error is dataitemcode or responsevalue
					if((subjDetail.SubjectResponseStatus == BufferAPI.BufferResponseStatus.FailDataItemCode)||
						(subjDetail.SubjectResponseStatus == BufferAPI.BufferResponseStatus.FailResponseValue))
					{
						// loop through reponses and show errors
						foreach(ResponseDetails respDet in subjDetail.Responses)
						{
							switch(respDet.BufferResponseStatus)
							{
								case BufferAPI.BufferResponseStatus.FailDataItemCode:
								case BufferAPI.BufferResponseStatus.FailResponseValue:
								{
									oXmlTw.WriteStartElement("MessageError");
									oXmlTw.WriteAttributeString("BufferErrorStatus", ((int)respDet.BufferResponseStatus).ToString());
									string sErrorRespDesc = ErrorDescription(subjDetail.SubjectResponseStatus);
									oXmlTw.WriteAttributeString("BufferErrorDescription", sErrorRespDesc);
									// extra info - not documented yet
									oXmlTw.WriteStartElement("ResponseDetails");
									oXmlTw.WriteAttributeString("Study", subjDetail.ClinicalTrialName);
									oXmlTw.WriteAttributeString("Site", subjDetail.Site);
									oXmlTw.WriteAttributeString("SubjectLabel", subjDetail.SubjectLabel);
									oXmlTw.WriteAttributeString("VisitCode", respDet.VisitCode);
									oXmlTw.WriteAttributeString("VisitCycle", respDet.VisitCycle.ToString());
									oXmlTw.WriteAttributeString("eFormCode", respDet.CRFPageCode);
									oXmlTw.WriteAttributeString("eFormCycle", respDet.CRFPageCycle.ToString());
									oXmlTw.WriteAttributeString("QuestionCode", respDet.DataItemCode);
									oXmlTw.WriteAttributeString("QuestionCycle", respDet.DataItemRepeatNo.ToString());
									oXmlTw.WriteAttributeString("QuestionValue", respDet.ResponseValue);

									// end responsedetails
									oXmlTw.WriteEndElement();
									// end messageerror
									oXmlTw.WriteEndElement();
									break;
								}
							}
						}
					}
				}

				// Close Message Errors element
				oXmlTw.WriteEndElement();

				// ResponseReport from attempt to write into MACRO
				oXmlTw.WriteStartElement("ResponseReport");
				oXmlTw.WriteStartElement("FailResponses");
				if(bMACROCommit)
				{
					// loop through reponses and show errors
					foreach(ResponseDetails respDet in subjDetail.Responses)
					{
						switch(respDet.BufferCommitStatus)
						{
							case BufferAPI.BufferCommitStatus.Success:
							{
								break;
							}
							default:
							{
								// all problems
								oXmlTw.WriteStartElement("FailResponse");
								oXmlTw.WriteAttributeString("FailStatus", ((int)respDet.BufferCommitStatus).ToString());
								oXmlTw.WriteAttributeString("FailDescription", CommitDescription(respDet.BufferCommitStatus));
								oXmlTw.WriteAttributeString("Study", subjDetail.ClinicalTrialName);
								oXmlTw.WriteAttributeString("Site", subjDetail.Site);
								oXmlTw.WriteAttributeString("SubjectLabel", subjDetail.SubjectLabel);
								oXmlTw.WriteAttributeString("VisitCode", respDet.VisitCode);
								oXmlTw.WriteAttributeString("VisitCycle", respDet.VisitCycle.ToString());
								oXmlTw.WriteAttributeString("eFormCode", respDet.CRFPageCode);
								oXmlTw.WriteAttributeString("eFormCycle", respDet.CRFPageCycle.ToString());
								oXmlTw.WriteAttributeString("QuestionCode", respDet.DataItemCode);
								oXmlTw.WriteAttributeString("QuestionCycle", respDet.CRFPageCycle.ToString());
								oXmlTw.WriteAttributeString("QuestionValue", respDet.ResponseValue);
								// Close FailResponses element
								oXmlTw.WriteEndElement();
								break;
							}
					
						}
					}
				}
				// close FailResponses
				oXmlTw.WriteEndElement();
				oXmlTw.WriteStartElement("FullReport");
				// write full report
				if(bMACROCommit)
				{
					oXmlTw.WriteStartElement("MACROSubject");
					oXmlTw.WriteAttributeString("Study", subjDetail.ClinicalTrialName);
					oXmlTw.WriteAttributeString("Site", subjDetail.Site);
					oXmlTw.WriteAttributeString("Label", subjDetail.SubjectLabel);
			
					// store previous visit / eform combination
					string sVisitPrev = "";
					int nVisitCyclePrev = BufferAPI._DEFAULT_MISSING_NUMERIC;
					string sEformPrev = "";
					int nEformPrevCycle = BufferAPI._DEFAULT_MISSING_NUMERIC;

					// loop through responses
					for(int i=0; i < subjDetail.Responses.Count; i++)
					{
						ResponseDetails respDet = (ResponseDetails)subjDetail.Responses[i];

						// check if new visit
						if((sVisitPrev != respDet.VisitCode)||(nVisitCyclePrev != respDet.VisitCycle))
						{
							// check if need to close eform & visit
							if(sEformPrev != "")
							{
								// close eform element & reset
								sEformPrev = "";
								oXmlTw.WriteEndElement();
							}
							if(sVisitPrev != "")
							{
								// close visit element
								oXmlTw.WriteEndElement();
							}
							// new visit
							oXmlTw.WriteStartElement("Visit");
							oXmlTw.WriteAttributeString("Code", respDet.VisitCode);
							oXmlTw.WriteAttributeString("Cycle", respDet.VisitCycle.ToString());
							// set prevs
							sVisitPrev = respDet.VisitCode;
							nVisitCyclePrev = respDet.VisitCycle;
						}

						// check if new eform
						if((sEformPrev != respDet.CRFPageCode)||(nEformPrevCycle != respDet.CRFPageCycle))
						{
							// check if need to close eform & visit
							if(sEformPrev != "")
							{
								// close eform element 
								oXmlTw.WriteEndElement();
							}
							// new eform element
							oXmlTw.WriteStartElement("eForm");
							oXmlTw.WriteAttributeString("Code", respDet.CRFPageCode);
							oXmlTw.WriteAttributeString("Cycle", respDet.CRFPageCycle.ToString());
							sEformPrev = respDet.CRFPageCode;
							nEformPrevCycle = respDet.CRFPageCycle;
						}

						// write response
						oXmlTw.WriteStartElement("Question");
						oXmlTw.WriteAttributeString("Code", respDet.DataItemCode);
						oXmlTw.WriteAttributeString("Cycle", respDet.DataItemRepeatNo.ToString());
						oXmlTw.WriteAttributeString("Value", respDet.ResponseValue);
						oXmlTw.WriteAttributeString("CommitStatus", ((int)respDet.BufferCommitStatus).ToString());
						// close Question element
						oXmlTw.WriteEndElement();

						// if last response
						if((i+1) == subjDetail.Responses.Count)
						{
							if(sEformPrev != "")
							{
								// close eform element & reset
								sEformPrev = "";
								oXmlTw.WriteEndElement();
							}
							if(sVisitPrev != "")
							{
								// close visit element
								oXmlTw.WriteEndElement();
							}
						}
					}

					// end MACROSubject
					oXmlTw.WriteEndElement();
				}
				// close FullReport
				oXmlTw.WriteEndElement();
				// Close ResponseReport
				oXmlTw.WriteEndElement();

				// end Message report
				oXmlTw.WriteEndElement();

				// end document
				oXmlTw.WriteEndDocument();
				oXmlTw.Flush();
				oXmlTw.Close();

				// collect xml string from memory stream
				ASCIIEncoding encoderAscii = new ASCIIEncoding();
				sReportXml = encoderAscii.GetString(oMemStream.ToArray());
			}
			catch(Exception ex)
			{	
				log.Error("Error creating report message", ex);
				sReportXml = "<BufferMessageReport BufferSaveStatus=\"8\"><MessageErrors><MessageError BufferErrorDescription=\"Error creating report message\" BufferErrorStatus=\"2\" /></MessageErrors></BufferMessageReport>";
			}
			return sReportXml;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="eBufferStatus"></param>
		/// <returns></returns>
		private string ErrorDescription(BufferAPI.BufferResponseStatus eBufferStatus)
		{
			string sErrorDesc = "";
			// switch buffer status
			switch(eBufferStatus)
			{
				case BufferAPI.BufferResponseStatus.FailParse:
				{
					sErrorDesc = "The supplied XML did not parse correctly.";
					break;
				}
				case BufferAPI.BufferResponseStatus.FailDbConnection:
				{
					sErrorDesc = "The connection to the MACRO database failed. Please check the settings and review the log file for more detail.";
					break;
				}
				case BufferAPI.BufferResponseStatus.FailClinicalTrial:
				{
					sErrorDesc = "The study code did not exist on the MACRO database.";
					break;
				}
				case BufferAPI.BufferResponseStatus.FailSite:
				{
					sErrorDesc = "The site code did not exist for the given study code on the MACRO database.";
					break;
				}
				case BufferAPI.BufferResponseStatus.FailSubjectLabel:
				{
					sErrorDesc = "The subject did not exist for the given study / site combination on the MACRO database.";
					break;
				}
				case BufferAPI.BufferResponseStatus.FailDataItemCode:
				{
					sErrorDesc = "A Data item did not exist for the given study code on the MACRO database.";
					break;
				}
				case BufferAPI.BufferResponseStatus.FailResponseValue:
				{
					sErrorDesc = "A response value did not exist / was invalid for the given study code on the MACRO database.";
					break;
				}
			}
			return sErrorDesc;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="eBufferStatus"></param>
		/// <returns></returns>
		private string CommitDescription(BufferAPI.BufferCommitStatus eBufferStatus)
		{
			string sMsgDesc = "";
			// switch buffer status
			switch(eBufferStatus)
			{
				case BufferAPI.BufferCommitStatus.EformInUse:
				{
					sMsgDesc = "eForm was in use by another user.";
					break;
				}
				case BufferAPI.BufferCommitStatus.EformLockedFrozen:
				{
					sMsgDesc = "eForm was locked or frozen.";
					break;
				}
				case BufferAPI.BufferCommitStatus.EformNotExist:
				{
					sMsgDesc = "eForm does not exist.";
					break;
				}
				case BufferAPI.BufferCommitStatus.InvalidXML:
				{
					sMsgDesc = "Invalid commit XML.";
					break;
				}
				case BufferAPI.BufferCommitStatus.LoginFailed:
				{
					sMsgDesc = "Login to MACRO database failed.";
					break;
				}
				case BufferAPI.BufferCommitStatus.NoLockForSave:
				{
					sMsgDesc = "No Lock obtained to commit save.";
					break;
				}
				case BufferAPI.BufferCommitStatus.NotCommitted:
				{
					sMsgDesc = "Not committed.";
					break;
				}
				case BufferAPI.BufferCommitStatus.NoVisitEformDate:
				{
					sMsgDesc = "No Visit/eForm date.";
					break;
				}
				case BufferAPI.BufferCommitStatus.QuestionNotEnterable:
				{
					sMsgDesc = "Question not enterable.";
					break;
				}
				case BufferAPI.BufferCommitStatus.QuestionNotExist:
				{
					sMsgDesc = "Question does not exist.";
					break;
				}
				case BufferAPI.BufferCommitStatus.SubjectNotExist:
				{
					sMsgDesc = "Subject does not exist.";
					break;
				}
				case BufferAPI.BufferCommitStatus.SubjectNotLoaded:
				{
					sMsgDesc = "Subject not loaded.";
					break;
				}
				case BufferAPI.BufferCommitStatus.Success:
				{
					sMsgDesc = "Success.";
					break;
				}
				case BufferAPI.BufferCommitStatus.ValueRejected:
				{
					sMsgDesc = "Value Rejected.";
					break;
				}
				case BufferAPI.BufferCommitStatus.VisitLockedFrozen:
				{
					sMsgDesc = "Visit was locked or frozen.";
					break;
				}
				case BufferAPI.BufferCommitStatus.VisitNotExist:
				{
					sMsgDesc = "Visit does not exist.";
					break;
				}
			}
			return sMsgDesc;
		}

	}
}
