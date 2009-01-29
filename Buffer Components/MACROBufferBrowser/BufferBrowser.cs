using System;
using log4net;
using InferMed.Components;
using System.Text;

namespace InferMed.MACROBuffer
{
	/// <summary>
	/// Summary description for BufferBrowser.
	/// </summary>
	class BufferBrowser
	{
		private BufferBrowser()
		{}

		// log4net
		private static readonly ILog log = LogManager.GetLogger( typeof(BufferBrowser) );

		public const int _DEFAULT_MISSING_NUMERIC = -1;
		public const string _USER_SETTINGS_FILE = "settingsfile";
		public const string _MACRO_CONN_DB = "bufferdb";
		public const string _BUFFER_TARGET_RANGE = "buffertargetrange";
		private const string _DEF_USER_SETTINGS_FILE = @"C:\Program Files\InferMed\MACRO 3.0\MACROUserSettings30.txt";
		private const string _SETTINGS_FILE = @"MACROSettings30.txt";
		public const string _SETTING_PAGE_LENGTH = "pagelength";
		public const string _DISPLAY_DATA_SQL_CLAUSE = "NOT IN ( 0, 16 ) ";

		public enum BufferCommitStatus
		{
			Success = 0,
			InvalidXML = 1,
			SubjectNotExist = 2,
			SubjectNotLoaded = 3,
			VisitNotExist = 4,
			EformNotExist = 5,
			QuestionNotExist = 6,
			EformInUse = 7,
			VisitLockedFrozen = 8,
			EformLockedFrozen = 9,
			QuestionNotEnterable = 10,
			NoVisitEformDate = 11,
			NoLockForSave = 12,
			ValueRejected = 13,
			NotCommitted = 14,
			LoginFailed = 15,
			DiscardedByUser = 16
		};

		/// <summary>
		/// Returns the MACRO usersettings file path from the settings file
		/// </summary>
		/// <returns>User settings file path</returns>
		public static string GetUserSettingsFilePath()
		{
			IMEDSettings20 iset = new IMEDSettings20( BufferBrowser._SETTINGS_FILE );
			return( iset.GetKeyValue( BufferBrowser._USER_SETTINGS_FILE, BufferBrowser._DEF_USER_SETTINGS_FILE ) );
		}

		/// <summary>
		/// Returns a setting from the MACRO 3.0 user settings file
		/// </summary>
		/// <param name="key">Setting key name</param>
		/// <param name="defaultVal">Default value to be returned if the key is not found</param>
		/// <returns>MACRO setting value</returns>
		public static string GetSetting(string key, string defaultVal)
		{
			IMEDSettings20 iset = new IMEDSettings20( GetUserSettingsFilePath() );
			return( iset.GetKeyValue( key, defaultVal ) );
		}

		/// <summary>
		/// Gets local Timezone in MACRO format
		/// </summary>
		/// <returns></returns>
		public static int GetLocalMACROTimezone()
		{
			// get local timezone
			TimeZone tzLocal = TimeZone.CurrentTimeZone;
			// get timespan of UTC offset
			TimeSpan tsLocal = tzLocal.GetUtcOffset(DateTime.Now);
			// calculate offset MACRO style - offset in minutes UTC (GMT)
			int nOffset = -1 * Convert.ToInt32(tsLocal.TotalMinutes);
			return nOffset;
		}

		/// <summary>
		/// return subject label given all info
		/// </summary>
		/// <param name="study"></param>
		/// <param name="site"></param>
		/// <param name="label"></param>
		/// <param name="subjectId"></param>
		/// <returns></returns>
		public static string SubjectLabel(string study, string site, string label, int subjectId)
		{
			StringBuilder sbLabel = new StringBuilder();

			sbLabel.Append( study );
			sbLabel.Append( @"\" );
			sbLabel.Append( site );
			sbLabel.Append( @"\" );
			// check if label is numeric
			if( label == subjectId.ToString() )
			{
				// numeric so use subject id
				sbLabel.Append( "(" );
				sbLabel.Append( subjectId.ToString() );
				sbLabel.Append( ")" );
			}
			else
			{
				// use label
				sbLabel.Append( label );
			}

			return sbLabel.ToString();
		}

		/// <summary>
		/// Calculate response status icon given response / lock / sdv / disc / note status
		/// </summary>
		/// <param name="responseStatus"></param>
		/// <param name="lockStatus"></param>
		/// <param name="sdvStatus"></param>
		/// <param name="discStatus"></param>
		/// <param name="note"></param>
		/// <param name="comment"></param>
		/// <returns></returns>
		public static string ResponseStatusIcon(int responseStatus, int lockStatus, int sdvStatus, int discStatus,
			bool note, bool comment)
		{
			string sHtml = "";
			string sCommentHtml = "";
			string sNoteHtml = "";
			string sStatusHtml = "";
			string sSDVHtml = "";
			string sSLabel = "";
			string sLLabel = "";
			string sSDVLabel = "";
			string sDLabel = "";
			string sToolTip = "";

			// note icon
			if(note)
			{
				sNoteHtml = @"<img src='../img/ico_note.gif'>";
			}

			// comment icon
			if( comment )
			{
				sCommentHtml = @"<img src='../img/ico_comment.gif'>";
			}

			// lock icon 
			switch( lockStatus )
			{
				case 6:
				{
					// frozen
					sStatusHtml = "ico_frozen";
					sLLabel = "Frozen";
					break;
				}
				case 5:
				{
					// locked
					sStatusHtml = "ico_locked";
					sLLabel = "Locked";
					break;
				}
			}

			// discrepancy icon
			switch ( discStatus )
			{
				case 20:
				{
					// responded
					if(sStatusHtml == "") 
					{
						sStatusHtml = "ico_disc_resp";
					}
					sDLabel = "Responded Discrepancy";
					break;
				}
				case 30:
				{
					// raised
					if(sStatusHtml == "") 
					{
						sStatusHtml = "ico_disc_raise";
					}
					sDLabel = "Raised Discrepancy";
					break;
				}
			}

			// status icon
			switch( responseStatus )
			{
				case -10:
					return( "" ); //requested
				case -20:
					// cancelled
					return( "Cancelled By User" );
				case -8:
				{
					// not applicable
					if(sStatusHtml == "") 
					{
						sStatusHtml = "ico_na";
					}
					sSLabel = "Not Applicable" + sSLabel;
					break;
				}
				case -5:
				{
					// unobtainable
					if( sStatusHtml == "" ) 
					{
						sStatusHtml = "ico_uo";
					}
					sSLabel = "Unobtainable" + sSLabel;
					break;
				}
				case 0:
				{
					// ok
					if(sStatusHtml == "") 
					{
						sStatusHtml = "ico_ok";
					}
					sSLabel = "OK" + sSLabel;
					break;
				}
				case 10:
				{
					// missing
					if(sStatusHtml == "") 
					{
						sStatusHtml = "ico_missing";
					}
					sSLabel = "Missing" + sSLabel;
					break;
				}
				case 20:
				{
					// inform
					if(sStatusHtml == "") 
					{
						sStatusHtml = "ico_inform";
					}
					sSLabel = "Inform" + sSLabel;
					break;
				}
				case 25:
				{
					// ok warning
					if(sStatusHtml == "") 
					{
						sStatusHtml = "ico_ok_warn";
					}
					sSLabel = "OK Warning" + sSLabel;
					break;
				}
				case 30:
				{
					// warning
					if(sStatusHtml == "" ) 
					{
						sStatusHtml = "ico_warn";
					}
					sSLabel = "Warning" + sSLabel;
					break;
				}
				case 40:
					break;
				default:
					break;
			}

			switch( sdvStatus )
			{
				case 20:
				{
					// complete
					sSDVHtml = @"<img src='../img/icof_sdv_done.gif'>";
					sSDVLabel = "Done SDV";
					break;
				}
				case 30:
				{
					// planned
					sSDVHtml = @"<img src='../img/icof_sdv_plan.gif'>";
					sSDVLabel = "Planned SDV";
					break;
				}
				case 40:
				{
					// queried
					sSDVHtml = @"<img src='../img/icof_sdv_query.gif'>";
					sSDVLabel = "Queried SDV";
					break;
				}
			}

			// tooltip
			if(sLLabel != "") 
			{
				sToolTip = sLLabel + ", ";
			}
			sToolTip = sToolTip + sSLabel;
			if(sSDVLabel != "")
			{
				sToolTip = sToolTip + ", " + sSDVLabel;
			}
			if(sDLabel != "")
			{
				sToolTip = sToolTip + ", " + sDLabel;
			}

			if(sStatusHtml != "")
			{
				sStatusHtml = "<img alt='" + sToolTip + @"' src='../img/" + sStatusHtml + ".gif'>";
			}

			// build image structure
			if((sNoteHtml == "") && (sCommentHtml == "") && (sSDVHtml == ""))
			{
				sHtml = sStatusHtml;
			}
			else
			{
				sHtml = "<table cellpadding='0' cellspacing='0'><tr>";
        
				if((sCommentHtml != "") || (sNoteHtml != ""))
				{
					sHtml = sHtml + "<td>";
					if((sCommentHtml != "") && (sNoteHtml != ""))
					{
						sHtml = sHtml + "<table cellpadding='0' cellspacing='0'>";
						sHtml = sHtml + "<tr><td>" + sCommentHtml + "</td></tr>";
						sHtml = sHtml + "<tr><td>" + sNoteHtml + "</td></tr>";
						sHtml = sHtml + "</table>";
					}
					else
					{
						sHtml = sHtml + sCommentHtml + sNoteHtml;
					}
					sHtml = sHtml + "</td>";
				}
        
				sHtml = sHtml + "<td>";
				if(sSDVHtml != "")
				{
					sHtml = sHtml + "<table cellpadding='0' cellspacing='0'>";
					sHtml = sHtml + "<tr><td>" + sStatusHtml + "</td></tr>";
					sHtml = sHtml + "<tr><td>" + sSDVHtml + "</td></tr>";
					sHtml = sHtml + "</table>";
				}
				else
				{
					sHtml = sHtml + sStatusHtml;
				}
				sHtml = sHtml + "</td>";
				sHtml = sHtml + "</tr></table>";
			}
			
			return sHtml;
		}

		/// <summary>
		/// Replace the html hex characters in the passed string
		/// </summary>
		/// <param name="html"></param>
		/// <returns></returns>
		public static string ReplaceHtmlCharacters(string html)
		{
			log.Info( "Starting ReplaceHtmlCharacters" );
			// replace spaces
			string htmlChars = html.Replace( "+", " ");
			log.Debug( "htmlChars = " + htmlChars );
			// hex decode chars
			int nPos = htmlChars.LastIndexOf( "%" );
			
			while( nPos > 0 )
			{
				string s1 = htmlChars.Substring(0, nPos);
				string s2 = ((char)System.Int32.Parse( htmlChars.Substring( nPos + 1, 2), System.Globalization.NumberStyles.AllowHexSpecifier )).ToString();
				string s3 = "";
				if( (nPos + 3) <= ( htmlChars.Length - 1 ) )
				{
					s3 = htmlChars.Substring( nPos + 3, (htmlChars.Length) - (nPos + 3));
				}
				// reform string
				htmlChars = s1 + s2 + s3;
				// NB: This takes the 2 characters after the % and converts them from hex to a single character.
				nPos = htmlChars.LastIndexOf( "%", nPos - 1 );
			}

			return htmlChars;
		}
	}
}
