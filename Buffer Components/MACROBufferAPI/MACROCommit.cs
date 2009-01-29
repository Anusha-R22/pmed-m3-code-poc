using System;
using log4net;
using InferMed.MACRO.API;
using System.Collections;
using System.Xml;
using System.Xml.Serialization;
using System.IO;

namespace InferMed.MACROBuffer
{
	/// <summary>
	/// Summary description for MACROCommit.
	/// </summary>
	class MACROCommit
	{
		private string _sMACROUser;
		private string _sMACROPassword;
		private string _sMACRODatabase;
		private string _sMACRORole;

		// log4net
		private static readonly ILog log = LogManager.GetLogger( typeof(MACROBufferAPI) );

		/// <summary>
		/// 
		/// </summary>
		/// <param name="sMACRODb"></param>
		public MACROCommit(string sMACRODb)
		{
			_sMACROUser = "bufferapi";
			_sMACROPassword = "m4cr04p1";
			_sMACRODatabase = sMACRODb;
			_sMACRORole = "MACROUser";
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="subjDetail"></param>
		/// <param name="bufferDb"></param>
		public void WriteToMACRO(ref SubjectDetails subjDetail, BufferResponseDb bufferDb)
		{
			// check subject can be committed to database - i.e. success parse/checks
			if(subjDetail.SubjectResponseStatus == BufferAPI.BufferResponseStatus.Success)
			{
				// create XML to write
				MACROSubject macroSubject = ConvertToSubjectData(ref subjDetail);
				string sXmlToCommit = SubjectToString(macroSubject);

				// commit to MACRO via InputXMLSubjectData
				string sCommitMessage = "";
				if(!CommitData( ref subjDetail, sXmlToCommit, ref sCommitMessage ))
				{
					// loop through failures and set as such on response objects
					ProcessAPIResponseString( ref subjDetail, sCommitMessage );
				}

				// update BufferResponseData / BufferResponseDataHistory
				bufferDb.UpdateBufferResponses( ref subjDetail );
			}
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="subjDetail"></param>
		/// <param name="sXmlToCommit"></param>
		/// <param name="sReturnMessage"></param>
		/// <returns></returns>
		private bool CommitData(ref SubjectDetails subjDetail, string sXmlToCommit, ref string sReturnMessage)
		{
			string sUsernameFull="";
			string sSerialisedUser="";

			// login
			log.Info("Log in to MACRO through COM API");
			API.LoginResult result = API.Login(_sMACROUser, _sMACROPassword, _sMACRODatabase, 
				_sMACRORole, ref sReturnMessage, ref sUsernameFull, ref sSerialisedUser);
			log.Info("Login Result - " + result.ToString() + ", user = " + _sMACROUser + ", MACRODb = " 
						+ _sMACRODatabase + ", MACRORole = " + _sMACRORole + " Message = " + sReturnMessage);
			// if successful
			if (result==API.LoginResult.Success )
			{
				// set username on responses
				SetAllUserDetails( ref subjDetail, _sMACROUser, sUsernameFull );

				// if login is a success commit data
				API.DataInputResult inputResult = API.InputXMLSubjectData(sSerialisedUser,sXmlToCommit,ref sReturnMessage);
				log.Info("Committing data to MACRO through COM API - " + Convert.ToInt16(inputResult));
				// if a success
				if (inputResult==API.DataInputResult.Success )
				{
					// all success
					SetAllResponseStatuses ( ref subjDetail, BufferAPI.BufferCommitStatus.Success );
					return true;
				} 
				else 
				{
					log.Warn(sReturnMessage);
					return false;
				}
			}
			else
			{
				// login failed
				log.Warn(sReturnMessage);
				sReturnMessage = "";
				SetAllResponseStatuses ( ref subjDetail, BufferAPI.BufferCommitStatus.LoginFailed );
				return false;
			}
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="subjDetail"></param>
		/// <param name="eCommitStatus"></param>
		private void SetAllResponseStatuses(ref SubjectDetails subjDetail, BufferAPI.BufferCommitStatus eCommitStatus)
		{
			// loop through responses
			foreach( ResponseDetails respDetail in subjDetail.Responses )
			{
				respDetail.BufferCommitStatus = eCommitStatus;
			}
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="subjDetail"></param>
		/// <param name="sUser"></param>
		/// <param name="sUsernameFull"></param>
		private void SetAllUserDetails(ref SubjectDetails subjDetail, string sUser, string sUsernameFull)
		{
			// loop through responses
			foreach( ResponseDetails respDetail in subjDetail.Responses )
			{
				respDetail.Username = sUser;
				respDetail.UsernameFull = sUsernameFull;
			}
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="subjDetail"></param>
		/// <returns></returns>
		private MACROSubject ConvertToSubjectData(ref SubjectDetails subjDetail)
		{
			MACROSubject subject = new MACROSubject();
			MACROSubjectVisit visit = null;

			// subject detail
			subject.Study=subjDetail.ClinicalTrialName;
			subject.Site=subjDetail.Site;
			subject.Label=subjDetail.SubjectLabel;

			// store previous visit / eform combination
			int nVisitPrev = BufferAPI._DEFAULT_MISSING_NUMERIC;
			int nVisitCyclePrev = BufferAPI._DEFAULT_MISSING_NUMERIC;

			// loop through responses
			for(int i=0; i < subjDetail.Responses.Count; i++)
			{
				ResponseDetails respDet = (ResponseDetails)subjDetail.Responses[i];

				// throw error if a required information piece is missing
				if((respDet.VisitId==BufferAPI._DEFAULT_MISSING_NUMERIC)||(respDet.VisitCycle==BufferAPI._DEFAULT_MISSING_NUMERIC)
					||(respDet.CRFPageId==BufferAPI._DEFAULT_MISSING_NUMERIC)||(respDet.CRFPageCycle==BufferAPI._DEFAULT_MISSING_NUMERIC)
					||(respDet.DataItemId==BufferAPI._DEFAULT_MISSING_NUMERIC)||(respDet.DataItemRepeatNo==BufferAPI._DEFAULT_MISSING_NUMERIC))
				{
					throw ( new Exception( "Missing Response Information." ) );
				}

				// check if new visit
				if((nVisitPrev != respDet.VisitId)||(nVisitCyclePrev != respDet.VisitCycle))
				{
					// create new visit
					visit = new MACROSubjectVisit();
					// add to subject
					subject.Visit.Add(visit);
					visit.Code = respDet.VisitCode;
					visit.Cycle = Convert.ToString(respDet.VisitCycle);
					// set prevs
					nVisitPrev = respDet.VisitId;
					nVisitCyclePrev = respDet.VisitCycle;
				}
				
				// check if eform exists
				int nEformPos = BufferAPI._DEFAULT_MISSING_NUMERIC;
				for(int nEform=0; nEform < visit.Eform.Count; nEform++)
				{
					MACROSubjectVisitEform visEform = (MACROSubjectVisitEform)visit.Eform[nEform];
					if(visEform != null)
					{
						if((visEform.Code == respDet.CRFPageCode)&&(visEform.Cycle == respDet.CRFPageCycle.ToString()))
						{
							nEformPos = nEform;
							break;
						}
					}
				}

				// get eform from object (if exists) or create new one
				MACROSubjectVisitEform eform = null;
				if(nEformPos != BufferAPI._DEFAULT_MISSING_NUMERIC)
				{
					eform = (MACROSubjectVisitEform)visit.Eform[nEformPos];
				}
				if (eform == null)
				{
					eform = new MACROSubjectVisitEform();
					eform.Code=respDet.CRFPageCode;
					eform.Cycle=Convert.ToString(respDet.CRFPageCycle);
					visit.Eform.Add(eform);
				}

				// add response to eform
				MACROSubjectVisitEformQuestion qu = new MACROSubjectVisitEformQuestion();
				qu.Code=respDet.DataItemCode;
				qu.Cycle=Convert.ToString(respDet.DataItemRepeatNo);
				qu.Value =respDet.ResponseValue;
				eform.Question.Add(qu);
			}
			return subject;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="subject"></param>
		/// <returns></returns>
		private static string SubjectToString(MACROSubject subject)
		{
			XmlSerializer x = new XmlSerializer(typeof(MACROSubject));
			MemoryStream stream = new MemoryStream();
			TextWriter writer = new StreamWriter(stream);;
			x.Serialize(writer, subject);
			writer.Close();
			string xml=new System.Text.UTF8Encoding().GetString(stream.ToArray());
			log.Debug(xml);
			return xml;
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="subjDetail"></param>
		/// <param name="sXmlCommitMsg"></param>
		private static void ProcessAPIResponseString(ref SubjectDetails subjDetail, string sXmlCommitMsg)
		{
			try
			{
				if(sXmlCommitMsg != "")
				{
					// load xml string
					XmlDocument commitStatusXML = new XmlDocument();
					commitStatusXML.LoadXml(sXmlCommitMsg);

					// loop through DataErrMsg
					foreach ( XmlNode nodedataErrMsg in commitStatusXML.SelectNodes( "//MACROInputErrors/DataErrMsg" ) )
					{
						// get detail from node
						// msgtype
						BufferAPI.BufferCommitStatus eCommitStatus = (BufferAPI.BufferCommitStatus)Convert.ToInt16(nodedataErrMsg.Attributes["MsgType"].Value);
						// description
						string sMsgDesc = nodedataErrMsg.Attributes["MsgDesc"].Value.ToString();
						// parse description
						if((eCommitStatus!=BufferAPI.BufferCommitStatus.InvalidXML)&&(eCommitStatus!=BufferAPI.BufferCommitStatus.SubjectNotLoaded)
							&&(eCommitStatus!=BufferAPI.BufferCommitStatus.Success))
						{
							// read to first space
							int nSpace = sMsgDesc.IndexOf(" ");
							if(nSpace > -1)
							{
								string sDetail = sMsgDesc.Substring(0, nSpace);
								// extract - should be in format visit[cycle]:eform[cycle]:question[cycle]
								string sMainSep=":";
								string[] asDetail = sDetail.Split(sMainSep.ToCharArray());
								string sVisitCode = "";
								int nVisitCycle = BufferAPI._DEFAULT_MISSING_NUMERIC;
								string sEFormCode = "";
								int nEFormCycle = BufferAPI._DEFAULT_MISSING_NUMERIC;
								string sQuestionCode = "";
								int nQuestionCycle = BufferAPI._DEFAULT_MISSING_NUMERIC;
								// get visit/eform/question detail
								switch(asDetail.Length)
								{
									case 1:
									{
										// visit
										GetCodeAndCycle(asDetail[0], ref sVisitCode, ref nVisitCycle);
										break;
									}
									case 2:
									{
										// visit
										GetCodeAndCycle(asDetail[0], ref sVisitCode, ref nVisitCycle);
										// eform
										GetCodeAndCycle(asDetail[1], ref sEFormCode, ref nEFormCycle);
										break;
									}
									case 3:
									{
										// visit
										GetCodeAndCycle(asDetail[0], ref sVisitCode, ref nVisitCycle);
										// eform
										GetCodeAndCycle(asDetail[1], ref sEFormCode, ref nEFormCycle);
										//question
										GetCodeAndCycle(asDetail[2], ref sQuestionCode, ref nQuestionCycle);
										break;
									}
								}
								// loop through responses
								for(int i=0; i < subjDetail.Responses.Count; i++)
								{
									ResponseDetails respDet = (ResponseDetails)subjDetail.Responses[i];
									if((sVisitCode == respDet.VisitCode)&&(nVisitCycle == respDet.VisitCycle))
									{
										if(sEFormCode != "")
										{
											if((sEFormCode == respDet.CRFPageCode)&&(nEFormCycle == respDet.CRFPageCycle))
											{
												if(sQuestionCode != "")
												{
													if((sQuestionCode == respDet.DataItemCode)&&(nQuestionCycle == respDet.DataItemRepeatNo))
													{
														// set status
														respDet.BufferCommitStatus = eCommitStatus;
													}
												}
												else
												{
													// set status
													respDet.BufferCommitStatus = eCommitStatus;
												}
											}
										}
										else
										{
											// set status of response
											respDet.BufferCommitStatus = eCommitStatus;
										}
									}
								}

							}
						}
					}
				}
			}
			catch(Exception ex)
			{
				log.Error( "Error reading commit response XML.", ex );
			}
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="sCodeAndCycle"></param>
		/// <param name="sCode"></param>
		/// <param name="sCycle"></param>
		private static void GetCodeAndCycle(string sCodeAndCycle, ref string sCode, ref int nCycle)
		{
			string sSepStart = "[";
			string sSepEnd = "]";
			string[] asCodeCycle = sCodeAndCycle.Split(sSepStart.ToCharArray());
			if( asCodeCycle.Length == 2)
			{
				// get code 
				sCode = asCodeCycle[0];
				// get cycle
				string sCycle = (asCodeCycle[1].Split(sSepEnd.ToCharArray()))[0];
				try
				{
					nCycle = Convert.ToInt16(sCycle);
				}
				catch
				{}
			}
		}
	}
}
