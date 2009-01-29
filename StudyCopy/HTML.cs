using System;
using System.Data;
using System.Text;
using log4net;

namespace InferMed.MACRO.StudyCopy
{
	/// <summary>
	/// html report generating code
	/// </summary>
	public class HTML
	{
		//logging
		private static readonly ILog log = LogManager.GetLogger( typeof( HTML ) );

		private HTML()
		{
		}

		/// <summary>
		/// Create an html report
		/// </summary>
		/// <param name="dbConn"></param>
		/// <param name="studyIdD"></param>
		/// <param name="studyCodeD"></param>
		/// <param name="studyIdS"></param>
		/// <param name="studyCodeS"></param>
		/// <param name="state"></param>
		/// <returns></returns>
		public static StringBuilder CreateReport( IDbConnection dbConn, string studyIdD, string studyCodeD, string studyIdS, 
			string studyCodeS, StudyState state, bool showEforms, bool showDataItems )
		{
			DataSet dsStudyD = new DataSet();
			DataSet dsStudyS = new DataSet();
			StringBuilder sbDocument = new StringBuilder();

			//logging
			log.Debug( "Creating report" );

			//open html document
			sbDocument.Append( "<HTML><HEAD><TITLE>SCT Study Differences Report</TITLE>" );
			
			sbDocument.Append( "<STYLE TYPE='text/css'> "
				+ "BODY,TABLE{FONT-FAMILY: verdana,arial,helvetica;FONT-SIZE: 8pt;COLOR: #000000;}"
				+ "</STYLE>" );

			sbDocument.Append( "</HEAD><BODY>" );

			//write report header
			sbDocument.Append( "SCT Study Differences Report " + DateTime.Now + " " 
				+ studyCodeD + " : " + studyCodeS +  "<BR><BR>" );

			//logging
			log.Debug( "Getting visit/eform structure" );

			//get study visit/eform structure
			dsStudyD = MACRO30.GetVisitEforms( dbConn, studyIdD );
			dsStudyS = MACRO30.GetVisitEforms( dbConn, studyIdS );

			//write eform and element table
			if( showEforms )
			{
				//logging
				log.Debug( "Creating eform table" );

				sbDocument.Append( CreateEformTable( dbConn, state, studyIdD, studyCodeD, studyIdS, studyCodeS, dsStudyD, dsStudyS ) );
				sbDocument.Append( "<BR><BR>" );
			}
			if( showDataItems )
			{
				//logging
				log.Debug( "Creating dataitem table" );

				sbDocument.Append( CreateDataItemTable( dbConn, state, studyIdD, studyCodeD, studyIdS, studyCodeS, dsStudyD ) );
			}

			//close html document
			sbDocument.Append( "</BODY></HTML>" );

			//return the report
			return( sbDocument );
		}

		/// <summary>
		/// Create an html eform table report
		/// </summary>
		/// <param name="dbConn"></param>
		/// <param name="state"></param>
		/// <param name="studyIdD"></param>
		/// <param name="studyCodeD"></param>
		/// <param name="studyIdS"></param>
		/// <param name="studyCodeS"></param>
		/// <param name="dsStudyD"></param>
		/// <param name="dsStudyS"></param>
		/// <returns></returns>
		private static string CreateEformTable( IDbConnection dbConn, StudyState state, string studyIdD, string studyCodeD, 
			string studyIdS, string studyCodeS, DataSet dsStudyD, DataSet dsStudyS )
		{
			StringBuilder sbTable = new StringBuilder();

			//write table header
			sbTable.Append( "<TABLE WIDTH='100%' BORDER='1' CELLPADDING='1' CELLSPACING='0'>" );
			sbTable.Append( "<TR HEIGHT='25'>" );
			sbTable.Append( "<TD ALIGN='CENTER' COLSPAN='2' WIDTH='45%'>" + studyCodeD + "</TD>" );
			sbTable.Append( "<TD ALIGN='CENTER' COLSPAN='2' WIDTH='45%'>" + studyCodeS + "</TD>" );
			sbTable.Append( "<TD ROWSPAN='2'WIDTH='10%'>Modification</TD>" );
			sbTable.Append( "</TR>" );
			sbTable.Append( "<TR HEIGHT='25'>" );
			sbTable.Append( "<TD WIDTH='20%'>Visit/eForm/Element</TD>" );
			sbTable.Append( "<TD WIDTH='25%'>Value</TD>" );
			sbTable.Append( "<TD WIDTH='20%'>Visit/eForm/Element</TD>" );
			sbTable.Append( "<TD WIDTH='25%'>Value</TD>" );
			sbTable.Append( "</TR>" );

			//loop through destination visit/eform
			foreach( DataRow rowVisitD in dsStudyD.Tables[0].Rows )
			{
				//get the state object for this eform
				Eform ef = state.GetEform( rowVisitD["CRFPAGEID"].ToString() );

				//get eform row for destination study
				DataSet dsD = MACRO30.GetEforms( dbConn, studyIdD, rowVisitD["CRFPAGEID"].ToString() );
				DataRow rowEformD = dsD.Tables[0].Rows[0];

				//create destination visit/eform caption
				string nameD = rowVisitD["VISITCODE"].ToString() + "/" + rowEformD["CRFTITLE"].ToString();
				string codeD = rowVisitD["VISITID"].ToString() + "/" + rowEformD["CRFPAGECODE"].ToString();

				//if the destination eform is matched we can compare its component parts
				if( ef.Matched )
				{
					//logging
					log.Debug( "Processing matched eform" + rowEformD["CRFPAGEID"].ToString() + " : " + ef.SourceId );

					//get eform for source study
					DataSet dsS = MACRO30.GetEforms( dbConn, studyIdS, ef.SourceId );
					DataRow rowEformS = dsS.Tables[0].Rows[0];
		
					//create source visit/eform caption
					string nameS, codeS;
					GetVisits( dsStudyS, ef.SourceId, out nameS, out codeS );
					nameS += "/" + rowEformS["CRFTITLE"].ToString();
					codeS += "/" + rowEformS["CRFPAGECODE"].ToString();

					//compare eform properties
					CompareRows( ref sbTable, rowEformD, rowEformS, codeD+"<BR>"+nameD, codeS+"<BR>"+nameS, "BACKGROUNDCOLOUR", 
						"CRFPAGELABEL", "LOCALCRFPAGELABEL", "DISPLAYNUMBERS", "HIDEIFINACTIVE", "EFORMWIDTH", "CRFTITLE" );
					
					//compare element properties
					CreateElementRows( dbConn, studyIdD, studyIdS, ref sbTable, codeD, nameD, codeS, nameS, ef, rowEformD["CRFPAGEID"].ToString(), 
						rowEformS["CRFPAGEID"].ToString() );

					//dispose of the dataset
					dsS.Dispose();
				}
				//the destination eform isnt matched so we assume it has been removed from the source study
				else
				{
					//logging
					log.Debug( "Processing unmatched eform " + rowEformD["CRFPAGEID"].ToString() );

					//add a REMOVED row to the table
					AddRowToHtml( ref sbTable, codeD+"<BR>"+nameD, "", "", "", "*EFORM REMOVED" );

					//add all of the elements on the eform to the table as they must also have been removed
					CreateAddedRemovedElementRows( dbConn, studyIdD, studyIdS, ref sbTable, false, codeD, nameD, 
						rowEformD["CRFPAGEID"].ToString() );
				}
				
				//dispose of the dataset
				dsD.Dispose();
			}


			//find all eforms that are in the source study but not matched in the destination study. these must have been added
			foreach( DataRow rowVisitS in dsStudyS.Tables[0].Rows )
			{
				if( !IsEformMatch( state, rowVisitS["CRFPAGEID"].ToString() ) )
				{
					//logging
					log.Debug( "Processing added eform" + rowVisitS["CRFPAGEID"].ToString() );

					//get eform for source study
					DataSet dsS = MACRO30.GetEforms( dbConn, studyIdS, rowVisitS["CRFPAGEID"].ToString() );
					DataRow rowEformS = dsS.Tables[0].Rows[0];

					//create source visit/eform caption
					string nameS, codeS;
					GetVisits( dsStudyS, rowVisitS["CRFPAGEID"].ToString(), out nameS, out codeS );
					nameS += "/" + rowEformS["CRFTITLE"].ToString();
					codeS += "/" + rowEformS["CRFPAGECODE"].ToString();

					//add an ADDED row to the table
					AddRowToHtml( ref sbTable, "", "", codeS+"<BR>"+nameS, "", "*EFORM ADDED" );

					//add all of the elements on the eform to the table as they must also have been added
					CreateAddedRemovedElementRows( dbConn, studyIdD, studyIdS, ref sbTable, true, codeS, nameS, rowEformS["CRFPAGEID"].ToString() );
				}
			}

			sbTable.Append( "</TABLE><BR><BR>" );
			return( sbTable.ToString() );
		}

		/// <summary>
		/// Create html rows of elements for added or removed eform
		/// </summary>
		/// <param name="dbConn"></param>
		/// <param name="studyIdD"></param>
		/// <param name="studyIdS"></param>
		/// <param name="sbTable"></param>
		/// <param name="added"></param>
		/// <param name="caption"></param>
		/// <param name="crfPageIdD"></param>
		private static void CreateAddedRemovedElementRows( IDbConnection dbConn, string studyIdD, string studyIdS, 
			ref StringBuilder sbTable, bool added, string codeCaption, string nameCaption, string crfPageIdD )
		{
			DataSet element = new DataSet();

			try
			{
				//get a list of eform elements
				element = MACRO30.GetEformElements( dbConn, ( ( added ) ? studyIdS : studyIdD ), crfPageIdD );

				//loop through the elements
				foreach( DataRow row in element.Tables[0].Rows )
				{
					//create caption
					string codeCap = codeCaption + "/" + row["CRFELEMENTID"].ToString();
					string nameCap = nameCaption + "/" + row["CAPTION"].ToString();

					//add element row
					if( added )
					{
						AddRowToHtml( ref sbTable, "", "", codeCap+"<BR>"+nameCap, "", "*ELEMENT ADDED" );
					}
					else
					{
						AddRowToHtml( ref sbTable, codeCap+"<BR>"+nameCap, "", "", "", "*ELEMENT REMOVED" );
					}
				}
			}
			finally
			{
				element.Dispose();
			}
		}

		/// <summary>
		/// Create html element rows comparing destination and source elements
		/// </summary>
		/// <param name="dbConn"></param>
		/// <param name="studyIdD"></param>
		/// <param name="studyIdS"></param>
		/// <param name="sbDocument"></param>
		/// <param name="codeD"></param>
		/// <param name="nameD"></param>
		/// <param name="codeS"></param>
		/// <param name="nameS"></param>
		/// <param name="ef"></param>
		/// <param name="crfPageIdD"></param>
		/// <param name="crfPageIdS"></param>
		private static void CreateElementRows( IDbConnection dbConn, string studyIdD, string studyIdS, ref StringBuilder sbDocument, 
			string codeD, string nameD, string codeS, string nameS, Eform ef, string crfPageIdD, string crfPageIdS )
		{
			DataSet elementD = new DataSet();
			DataSet elementS = new DataSet();

			try
			{
				//logging
				log.Debug( "Creating element rows " + crfPageIdD + " : " + crfPageIdS );

				//get elements for destination and source eforms
				elementD = MACRO30.GetEformElements( dbConn, studyIdD, crfPageIdD );
				elementS = MACRO30.GetEformElements( dbConn, studyIdS, crfPageIdS );

				//loop through destination elements
				foreach( DataRow rowD in elementD.Tables[0].Rows )
				{
					//get destination state element
					EformElement el = ef.GetElementByDestination( rowD["CRFELEMENTID"].ToString() );

					//create destination caption
					string cD = codeD + "/" + rowD["CRFELEMENTID"].ToString();
					string nD = nameD + "/" + rowD["CAPTION"].ToString();

					//if the element has a match, compare the match
					if( el.Matched )
					{
						//get a datarow for the source element
						DataRow rowS = GetElementRow( elementS, el.SourceId );

						//create the source caption
						string cS = codeS + "/" + rowS["CRFELEMENTID"].ToString();
						string nS = nameS + "/" + rowS["CAPTION"].ToString();

						//compare the element properties
						CompareRows( ref sbDocument, rowD, rowS, cD+"<BR>"+nD, cS+"<BR>"+nS,
							"CONTROLTYPE", "FONTCOLOUR", "CAPTION", "FONTNAME", "FONTBOLD", "FONTITALIC", "FONTSIZE", "FIELDORDER", 
							"SKIPCONDITION", "HEIGHT", "WIDTH", "CAPTIONX", "CAPTIONY", "X", "Y", "PRINTORDER", "HIDDEN", "LOCALFLAG",
							"OPTIONAL", "MANDATORY", "REQUIRECOMMENT", "ROLECODE", "QGROUPFIELDORDER", "SHOWSTATUSFLAG", 
							"CAPTIONFONTNAME", "CAPTIONFONTBOLD", "CAPTIONFONTITALIC", "CAPTIONFONTSIZE", "CAPTIONFONTCOLOUR", 
							"ELEMENTUSE", "DISPLAYLENGTH" );
					}
					//if the element has no match assume it has been removed from the source eform
					else
					{
						//add the element row
						AddRowToHtml( ref sbDocument, cD+"<BR>"+nD, "", codeS+"<BR>"+nameS, "", "*ELEMENT REMOVED" );
					}
				}

				//find all elements that are in the source eform but not matched in the destination eform
				foreach( DataRow rowS in elementS.Tables[0].Rows )
				{
					//if the sourec element hasnt been matched
					if( !IsElementMatch( ef, rowS["CRFELEMENTID"].ToString() ) )
					{
						//create the source caption
						string cS = codeS + "/" + rowS["CRFELEMENTID"].ToString();
						string nS = nameS + "/" + rowS["CAPTION"].ToString();

						//add the element row
						AddRowToHtml( ref sbDocument, codeD+"<BR>"+nameD, "", cS+"<BR>"+nS, "", "*ELEMENT ADDED" );
					}
				}


			}
			finally
			{
				elementD.Dispose();
				elementS.Dispose();
			}
		}

		/// <summary>
		/// Create html dataitem table
		/// </summary>
		/// <param name="dbConn"></param>
		/// <param name="state"></param>
		/// <param name="studyIdD"></param>
		/// <param name="studyCodeD"></param>
		/// <param name="studyIdS"></param>
		/// <param name="studyCodeS"></param>
		/// <param name="dsStudyD"></param>
		/// <returns></returns>
		private static string CreateDataItemTable( IDbConnection dbConn, StudyState state, string studyIdD, string studyCodeD, 
			string studyIdS, string studyCodeS, DataSet dsStudyD )
		{
			StringBuilder sbTable = new StringBuilder();

			//write table header
			sbTable.Append( "<TABLE WIDTH='100%' BORDER='1' CELLPADDING='1' CELLSPACING='0'>" );
			sbTable.Append( "<TR HEIGHT='25'>" );
			sbTable.Append( "<TD ALIGN='CENTER' COLSPAN='2' WIDTH='45%'>" + studyCodeD + "</TD>" );
			sbTable.Append( "<TD ALIGN='CENTER' COLSPAN='2' WIDTH='45%'>" + studyCodeS + "</TD>" );
			sbTable.Append( "<TD ROWSPAN='2'WIDTH='10%'>Modification</TD>" );
			sbTable.Append( "</TR>" );
			sbTable.Append( "<TR HEIGHT='25'>" );
			sbTable.Append( "<TD WIDTH='20%'>Visit/eForm/DataItem</TD>" );
			sbTable.Append( "<TD WIDTH='25%'>Value</TD>" );
			sbTable.Append( "<TD WIDTH='20%'>Visit/eForm/DataItem</TD>" );
			sbTable.Append( "<TD WIDTH='25%'>Value</TD>" );
			sbTable.Append( "</TR>" );

			//loop through visit/eforms
			foreach( DataRow rowVisitD in dsStudyD.Tables[0].Rows )
			{
				//get eform for destination study
				DataSet dsD = MACRO30.GetEforms( dbConn, studyIdD, rowVisitD["CRFPAGEID"].ToString() );
				DataRow rowEformD = dsD.Tables[0].Rows[0];
				Eform ef = state.GetEform( rowEformD["CRFPAGEID"].ToString() );
			
				//if the eform is matched
				if( ef.Matched )
				{
					//logging
					log.Debug( "Creating dataitem rows " + rowEformD["CRFPAGEID"].ToString() );

					//get source eform
					DataSet dsS = MACRO30.GetEforms( dbConn, studyIdS, ef.SourceId );
					DataRow rowEformS = dsS.Tables[0].Rows[0];

					//get elements for source and destination eforms
					DataSet elementD = MACRO30.GetEformElements( dbConn, studyIdD, rowEformD["CRFPAGEID"].ToString() );
					DataSet elementS = MACRO30.GetEformElements( dbConn, studyIdS, ef.SourceId );

					//loop through each destination element
					foreach( DataRow rowD in elementD.Tables[0].Rows )
					{
						EformElement el = ef.GetElementByDestination( rowD["CRFELEMENTID"].ToString() );
			
						//if the element is matched
						if( el.Matched )
						{
							//create the destination caption
							string nameD = rowVisitD["VISITCODE"].ToString() + "/" + rowEformD["CRFTITLE"].ToString() 
								+ "/" + rowD["DATAITEMCODE"].ToString();
							string codeD = rowVisitD["VISITID"].ToString() + "/" + rowEformD["CRFPAGECODE"].ToString() + ")"
								+ "/" + rowD["DATAITEMID"].ToString();

							//get the source element row
							DataRow rowS = GetElementRow( elementS, el.SourceId );

							//compare the dataitem properties
							CompareRows( ref sbTable, rowD, rowS, codeD+"<BR>"+nameD, rowS["DATAITEMCODE"].ToString(), "DATATYPE", 
								"DATAITEMFORMAT", "UNITOFMEASUREMENT", "DATAITEMLENGTH","DERIVATION", "DATAITEMHELPTEXT", 
								"DATAITEMCASE", "DESCRIPTION" );
					
							//compare the validation rows
							CreateValidationRows(dbConn, studyIdD, studyIdS, ref sbTable, codeD+"<BR>"+nameD, rowS["DATAITEMCODE"].ToString(),
								rowD["DATAITEMID"].ToString(), rowS["DATAITEMID"].ToString() );

							//compare the category rows if any
							CreateCategoryRows( dbConn, studyIdD, studyIdS, ref sbTable, codeD+"<BR>"+nameD, rowS["DATAITEMCODE"].ToString(),
								rowD, rowS );
						}
					}
					dsS.Dispose();
				}
				dsD.Dispose();
			}

			sbTable.Append( "</TABLE><BR><BR>" );
			return( sbTable.ToString() );
		}

		/// <summary>
		/// Create html category rows
		/// </summary>
		/// <param name="dbConn"></param>
		/// <param name="studyIdD"></param>
		/// <param name="studyIdS"></param>
		/// <param name="sbTable"></param>
		/// <param name="captionD"></param>
		/// <param name="captionS"></param>
		/// <param name="rowD"></param>
		/// <param name="rowS"></param>
		private static void CreateCategoryRows( IDbConnection dbConn, string studyIdD, string studyIdS, 
			ref StringBuilder sbTable, string captionD, string captionS, DataRow rowD, DataRow rowS )
		{
			DataSet catValuesD = new DataSet();
			DataSet catValuesS = new DataSet();

			try
			{
				//logging
				log.Debug( "Creating category rows " + rowD["DATAITEMID"].ToString() +  " : " + rowS["DATAITEMID"].ToString() );

				//get category values for source and destination
				catValuesD = MACRO30.GetCategoryValues( dbConn, studyIdD, rowD );
				catValuesS = MACRO30.GetCategoryValues( dbConn, studyIdS, rowS );

				//check whether source or destination has more value rows
				int maxD = catValuesD.Tables[0].Rows.Count;
				int maxS = catValuesS.Tables[0].Rows.Count;
				int max = ( maxD > maxS )? maxS : maxD;
					
				//compare all rows
				for( int n = 0; n < max; n++ )
				{
					CompareRows( ref sbTable, catValuesD.Tables[0].Rows[n], catValuesS.Tables[0].Rows[n], 
						captionD, captionS, "VALUEID", "VALUECODE", "ITEMVALUE", "ACTIVE", "VALUEORDER" );
				}

				//add a row for removed category values
				if( maxD > maxS )
				{
					for( int n = 0; n < maxD; n++ )
					{
						AddRowToHtml( ref sbTable, captionD, catValuesD.Tables[0].Rows[n]["VALUECODE"].ToString(), 
							captionS, "", "* CATEGORY REMOVED" );
					}
				}
				//add a row for added category values
				else if ( maxS > maxD )
				{
					for( int n = 0; n < maxS; n++ )
					{
						AddRowToHtml( ref sbTable, captionD, "", captionS, catValuesS.Tables[0].Rows[n]["VALUECODE"].ToString(), 
							"*CATEGORY ADDED" );
					}
				}
			}
			finally
			{
				catValuesD.Dispose();
				catValuesS.Dispose();
			}
		}

		/// <summary>
		/// Create html validation rows
		/// </summary>
		/// <param name="dbConn"></param>
		/// <param name="studyIdD"></param>
		/// <param name="studyIdS"></param>
		/// <param name="sbTable"></param>
		/// <param name="captionD"></param>
		/// <param name="captionS"></param>
		/// <param name="dataItemIdD"></param>
		/// <param name="dataItemIdS"></param>
		private static void CreateValidationRows( IDbConnection dbConn, string studyIdD, string studyIdS, 
			ref StringBuilder sbTable, string captionD, string captionS, string dataItemIdD, string dataItemIdS )
		{
			DataSet validationsD = new DataSet();
			DataSet validationsS = new DataSet();

			try
			{
				//logging
				log.Debug( "Creating validation rows " + dataItemIdD +  " : " + dataItemIdS );

				//get validations for source and destination
				validationsD = MACRO30.GetDataItemValidation( dbConn, studyIdD, dataItemIdD );
				validationsS = MACRO30.GetDataItemValidation( dbConn, studyIdS, dataItemIdS );

				//check whether source or destination has more validations
				int maxD = validationsD.Tables[0].Rows.Count;
				int maxS = validationsS.Tables[0].Rows.Count;
				int max = ( maxD > maxS )? maxS : maxD;
				
				//compare all rows
				for( int n = 0; n < max; n++ )
				{
					CompareRows( ref sbTable, validationsD.Tables[0].Rows[n], validationsS.Tables[0].Rows[n], 
						captionD, captionS, "VALIDATIONTYPEID", "DATAITEMVALIDATION", "VALIDATIONMESSAGE" );
				}

				//add a row for removed validations
				if( maxD > maxS )
				{
					for( int n = 0; n < maxD; n++ )
					{
						AddRowToHtml( ref sbTable, captionD, validationsD.Tables[0].Rows[n]["DATAITEMVALIDATION"].ToString(), 
							captionS, "", "* VALIDATION REMOVED" );
					}
				}
				//add a row for added validations
				else if ( maxS > maxD )
				{
					for( int n = 0; n < maxS; n++ )
					{
						AddRowToHtml( ref sbTable, captionD, "", captionS, validationsS.Tables[0].Rows[n]["DATAITEMVALIDATION"].ToString(), 
							"*VALIDATION ADDED" );
					}
				}
			}
			finally
			{
				validationsD.Dispose();
				validationsS.Dispose();
			}
		}

		/// <summary>
		/// Return an element row from a dataset matched on Id
		/// </summary>
		/// <param name="el"></param>
		/// <param name="crfElementId"></param>
		/// <returns></returns>
		private static DataRow GetElementRow( DataSet el, string crfElementId )
		{
			DataRow match = null;
			int n = 0;

			while( ( n < el.Tables[0].Rows.Count ) && ( match == null ) )
			{
				DataRow row = el.Tables[0].Rows[n];
				if( row["CRFELEMENTID"].ToString() == crfElementId ) match = row;
				n++;
			}

			return( match );
		}

		/// <summary>
		/// Is a source element matched
		/// </summary>
		/// <param name="ef"></param>
		/// <param name="elementId"></param>
		/// <returns></returns>
		private static bool IsElementMatch( Eform ef, string elementId )
		{
			foreach( EformElement el in ef.Elements )
			{
				if( el.SourceId == elementId ) return( true );
			}
			return( false );
		}

		/// <summary>
		/// Is source eform matched
		/// </summary>
		/// <param name="state"></param>
		/// <param name="eformId"></param>
		/// <returns></returns>
		private static bool IsEformMatch( StudyState state, string eformId )
		{
			foreach( Eform ef in state.Eforms )
			{
				if( ef.SourceId == eformId ) return( true );
			}
			return( false );
		}

		/// <summary>
		/// Get a delimited list of visits to which an eform belongs
		/// </summary>
		/// <param name="vi"></param>
		/// <param name="eformId"></param>
		/// <param name="nameCaption"></param>
		/// <param name="codeCaption"></param>
		private static void GetVisits( DataSet vi, string eformId, out string nameCaption, out string codeCaption )
		{
			nameCaption = ""; 
			codeCaption = "";

			foreach( DataRow row in vi.Tables[0].Rows )
			{
				if( row["CRFPAGEID"].ToString() == eformId ) 
				{
					nameCaption += row["VISITCODE"].ToString() + ",";
					codeCaption += row["VISITID"].ToString() + ",";
				}
			}
			if( nameCaption.EndsWith( "," ) ) nameCaption = ( nameCaption.Substring( 0, nameCaption.Length - 1 ) );
			if( codeCaption.EndsWith( "," ) ) codeCaption = ( codeCaption.Substring( 0, codeCaption.Length - 1 ) );
		}

		/// <summary>
		/// Compare a list of parameters in 2 datarows
		/// </summary>
		/// <param name="sb"></param>
		/// <param name="rowD"></param>
		/// <param name="rowS"></param>
		/// <param name="captionD"></param>
		/// <param name="captionS"></param>
		/// <param name="cols"></param>
		private static void CompareRows( ref StringBuilder sb, DataRow rowD, DataRow rowS, string captionD, string captionS, 
			params string[] cols )
		{
			foreach( string col in cols )
			{
				if( !ValuesMatch( rowD, rowS, col ) )
				{
					AddRowToHtml( ref sb, captionD, rowD[col].ToString(), captionS, rowS[col].ToString(), col );
				}
			}
		}

		/// <summary>
		/// Compare 2 datarow values
		/// </summary>
		/// <param name="rowD"></param>
		/// <param name="rowS"></param>
		/// <param name="col"></param>
		/// <returns></returns>
		private static bool ValuesMatch( DataRow rowD, DataRow rowS, string col )
		{
			return( rowD[col].ToString() == rowS[col].ToString() );
		}

		/// <summary>
		/// Add a row to an html table
		/// </summary>
		/// <param name="sb"></param>
		/// <param name="captionD"></param>
		/// <param name="valueD"></param>
		/// <param name="captionS"></param>
		/// <param name="valueS"></param>
		/// <param name="modification"></param>
		private static void AddRowToHtml( ref StringBuilder sb, string captionD, string valueD, string captionS, string valueS,
			string modification )
		{
			sb.Append( "<TR VALIGN='TOP'>" );
			sb.Append( "<TD>" + captionD + "&nbsp;</TD>" );
			sb.Append( "<TD>" + valueD + "&nbsp;</TD>" );
			sb.Append( "<TD>" + captionS + "&nbsp;</TD>" );
			sb.Append( "<TD>" + valueS + "&nbsp;</TD>" );
			sb.Append( "<TD>" + modification + "&nbsp;</TD>" );
			sb.Append( "</TR>" );
		}
	}
}
