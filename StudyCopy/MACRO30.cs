using System;
using System.Data;
using System.Collections;
using System.Windows.Forms;
using System.Data.OracleClient;
using System.Text;
using InferMed.Components;
using log4net;


//----------------------------------------------------------------------
// 03/05/2006	issue 2729 enter description when adding dataitem
// 03/05/2006 bug 2730 check studyvisit, crfpage and dataitem for next available dataitemid
//----------------------------------------------------------------------

namespace InferMed.MACRO.StudyCopy
{
	/// <summary>
	/// MACRO30 study copying and data access definitions
	/// </summary>
	public class MACRO30
	{
		//callback delegate
		public delegate void DShowProgress(string sProg, bool bLog, StudyCopyGlobal.LogPriority pr);
		//callback delegate event handler
		public static event DShowProgress DShowProgressEvent;

		//logging
		private static readonly ILog log = LogManager.GetLogger( typeof( MACRO30 ) );

		//minimum dataitem id for new dataitem ids
		private const string _MINDATAITEMID = "20001";

		//minimum qgroupid for new qgroups
		private const string _MINQGROUPID = "28001";

		private MACRO30()
		{
		}

		/// <summary>
		/// Check for a permission belonging to a role
		/// </summary>
		/// <param name="dbConn"></param>
		/// <param name="roleCode"></param>
		/// <param name="functionCode"></param>
		/// <returns></returns>
		public static bool CheckPermission(IDbConnection dbConn, string roleCode, string functionCode)
		{
			DataSet ds = new DataSet();

			try
			{
				string sql = "SELECT COUNT(*) AS MATCHES "
					+ "FROM ROLEFUNCTION "
					+ "WHERE ROLECODE = '" + roleCode + "' "
					+ "AND FUNCTIONCODE = '" + functionCode + "' ";

				//execute the query
				ds = GetDataSet( dbConn, sql );
				//return whether the dataitemcode was found or not 
				return( ( System.Convert.ToInt32( ( ds.Tables[0].Rows[0]["MATCHES"].ToString() ) ) > 0 ) );
			}
			finally
			{
				ds.Dispose();
			}
		}

		/// <summary>
		/// Renumber all eform elements after adding a new element to an eform
		/// This is a straight copy of the VB Renumber() function in SD module
		/// </summary>
		/// <param name="dbConn"></param>
		/// <param name="studyId"></param>
		/// <param name="versionId"></param>
		/// <param name="eformId"></param>
		public static void RenumberEformElements(IDbConnection dbConn, string studyId, string versionId, string eformId )
		{
			DataSet ds = new DataSet();
			short fieldOrder = 0;

			try
			{
				ds = GetElementsForRenumber(dbConn, studyId, versionId, eformId );
 
				DShowProgressEvent("Renumbering eForm elements", false, StudyCopyGlobal.LogPriority.Low);
				
				foreach( DataRow row in ds.Tables[0].Rows )
				{
					string sql = "UPDATE CRFELEMENT ";
					string sqlAnd = "";
					
					if( ( row["HIDDEN"].ToString() != "1" ) && ( row["ELEMENTUSE"].ToString() == "0" ) )
					{
						fieldOrder++;
						sql += "SET FIELDORDER = " + fieldOrder + " ";
					}
					else
					{
						//hidden dataitem so set fieldorder to 0
						sql += "SET FIELDORDER = 0 ";
					}

					sql += "WHERE CLINICALTRIALID = " + studyId + " "
						+ "AND VERSIONID = " + versionId + " "
						+ "AND CRFPAGEID = " + eformId + " ";
					
					sqlAnd = "AND CRFELEMENTID = " + row["CRFELEMENTID"].ToString();

					//logging
					log.Debug( sql + sqlAnd );

					//update element
					RunSQL( dbConn, sql + sqlAnd );

					if( System.Convert.ToInt16( row["QGROUPID"].ToString() ) > 0 )
					{
						sqlAnd = "AND OWNERQGROUPID = " + row["QGROUPID"].ToString();

						//logging
						log.Debug( sql + sqlAnd );

						RunSQL( dbConn, sql + sqlAnd );
					}
				}
			}
			finally
			{
				ds.Dispose();
			}
		}

		public static void ShowProgress( string prog )
		{
			DShowProgressEvent( prog, false, StudyCopyGlobal.LogPriority.Low );
		}

		/// <summary>
		/// Get all the numbered elements on an eform in order of x, y
		/// </summary>
		/// <param name="dbConn"></param>
		/// <param name="studyId"></param>
		/// <param name="versionId"></param>
		/// <param name="eformId"></param>
		/// <returns></returns>
		public static DataSet GetElementsForRenumber(IDbConnection dbConn, string studyId, string versionId, string eformId )
		{
			DShowProgressEvent("Getting eFormElement data from database", false, StudyCopyGlobal.LogPriority.Low);
			string sql = "";

			sql = "SELECT * FROM CRFELEMENT "
				+ "WHERE CLINICALTRIALID = " + studyId + " "
				+ "AND VERSIONID = " + versionId + " "
        + "AND CRFPAGEID = " + eformId + " "
        + "AND ( DATAITEMID > 0 OR QGROUPID > 0 ) "
        + "AND ( OWNERQGROUPID = 0 ) "
        + "AND ((FIELDORDER > 0) OR (HIDDEN = 0) OR (ELEMENTUSE = 0)) "
        + "ORDER BY Y, X";

			//logging
			log.Debug( sql );

			//return the data set
			return (GetDataSet(dbConn, sql));
		}

		/// <summary>
		/// Copy the eform properties of a MACRO formatted study into the MACRO unformatted study
		/// </summary>
		/// <param name="dbConn"></param>
		/// <param name="studyIdS"></param>
		/// <param name="studyIdD"></param>
		/// <param name="rowS"></param>
		/// <param name="rowD"></param>
		public static void CopyEform(IDbConnection dbConn, string studyIdS, string studyIdD, DataRow rowS, DataRow rowD,
			ArrayList copyProperties )
		{
			//logging
			log.Debug( "Updating database" );

			DShowProgressEvent("Copying eForm attributes, " + rowD["CRFPAGEID"] + " : " + rowD["CRFTITLE"], false, StudyCopyGlobal.LogPriority.Normal);

			//build the update sql
			string sql = 
				"UPDATE CRFPAGE "
				+ "SET "
				+ GetSetPropertySql( rowS, copyProperties ) + " "
				+ "WHERE CLINICALTRIALID=" + studyIdD + " AND CRFPAGEID=" + rowD["CRFPAGEID"];

			//logging
			log.Info( "Copying eform:"
				+ " study:" + studyIdS + "->" + studyIdD
				+ " eform:" + rowS["CRFPAGEID"].ToString() + "->" + rowD["CRFPAGEID"].ToString() 
				+ " properties:" + GetCopyPropertyString( copyProperties ) );
			log.Debug( sql );

			//execute the update
			RunSQL( dbConn, sql );

			//logging
			log.Debug( "Database updated" );
		}

		/// <summary>
		/// Create sql for an arraylist of study properties
		/// </summary>
		/// <param name="rowS"></param>
		/// <param name="copyProperties"></param>
		/// <returns></returns>
		private static string GetSetPropertySql( DataRow rowS, ArrayList copyProperties )
		{
			string sql = "";

			for( int n = 0; n < copyProperties.Count; n++ )
			{
				sql += copyProperties[n].ToString() + "=" + GetDBValue( rowS, copyProperties[n].ToString() ) + ", ";
			}
			if( sql.EndsWith( ", " ) )
			{
				sql = sql.Substring( 0, ( sql.Length - 2 ) );
			}

			return( sql );
		}

		/// <summary>
		/// Copy the  eform element properties of a MACRO formatted study into a MACRO unformatted study
		/// </summary>
		/// <param name="dbConn"></param>
		/// <param name="studyIdD"></param>
		/// <param name="rowS"></param>
		/// <param name="rowD"></param>
		/// <returns></returns>
		public static void CopyEformElement(IDbConnection dbConn, string studyIdD, string studyIdS, DataRow rowS, DataRow rowD,
			ArrayList copyProperties )
		{
			//logging
			log.Debug( "Updating database" );

			DShowProgressEvent("Copying eFormElement attributes, " + rowD["CRFELEMENTID"], false, StudyCopyGlobal.LogPriority.Low);

			//build the update sql
			string sql =
				"UPDATE CRFELEMENT "
				+ "SET "
				+ GetSetPropertySql( rowS, copyProperties ) + " "
				+ "WHERE CLINICALTRIALID=" + studyIdD + " AND VERSIONID=" + rowD["VERSIONID"].ToString() + " "
				+ "AND CRFPAGEID=" + rowD["CRFPAGEID"].ToString() + " AND CRFELEMENTID=" + rowD["CRFELEMENTID"].ToString();

			//logging
			log.Info( "Copying element:"
				+ " study:" + studyIdS + "->" + studyIdD
				+ " eform:" + rowS["CRFPAGEID"].ToString() + "->" + rowD["CRFPAGEID"].ToString() 
				+ " element:" + rowS["CRFELEMENTID"].ToString() + "->" + rowD["CRFELEMENTID"].ToString() 
				+ " properties:" + GetCopyPropertyString( copyProperties ) );
			log.Debug( sql );

			//execute the update
			RunSQL( dbConn, sql );

			//logging
			log.Debug( "Database updated" );
		}

		/// <summary>
		/// Insert a element into an eForm
		/// </summary>
		/// <param name="dbConn"></param>
		/// <param name="rowS"></param>
		/// <param name="studyIdD"></param>
		/// <param name="studyIdS"></param>
		/// <param name="eFormIdD"></param>
		/// <param name="dataItemId"></param>
		/// <returns></returns>
		public static DataRow InsertEformElement( IDbConnection dbConn, DataRow rowS, string studyIdD, string studyIdS, string eFormIdD, 
			string dataItemId, string qGroupIdD )
		{
			string sql = "";
			DataSet ds = new DataSet();
			DataRow newRow = null;
			
			//logging
			log.Debug( "Updating database" );

			try
			{
				string elementId = GetNextEformElementID( dbConn, studyIdD, rowS["VERSIONID"].ToString(), eFormIdD);

				DShowProgressEvent("Inserting eFormElement," + rowS["CRFELEMENTID"].ToString(), false, StudyCopyGlobal.LogPriority.Low);

				//insert a new element
				sql = 
					"INSERT INTO CRFELEMENT "
					+ "(CLINICALTRIALID, CRFPAGEID, VERSIONID, CRFELEMENTID, DATAITEMID, CONTROLTYPE, FONTCOLOUR, CAPTION, "
					+ "FONTNAME, FONTBOLD, FONTITALIC, FONTSIZE, FIELDORDER, SKIPCONDITION, HEIGHT, WIDTH, CAPTIONX, "
					+ "CAPTIONY, X, Y, PRINTORDER, HIDDEN, LOCALFLAG, OPTIONAL, MANDATORY, REQUIRECOMMENT, ROLECODE, "
					+ "OWNERQGROUPID, QGROUPID, QGROUPFIELDORDER, SHOWSTATUSFLAG, CAPTIONFONTNAME, CAPTIONFONTBOLD, "
					+ "CAPTIONFONTITALIC, CAPTIONFONTSIZE, CAPTIONFONTCOLOUR, ELEMENTUSE, DISPLAYLENGTH, HOTLINK) "
					+ "VALUES (" + studyIdD + ", " + eFormIdD + ", " + GetDBValue(rowS, "VERSIONID") + ", "
					+ elementId + ", " + dataItemId + ", " + GetDBValue(rowS, "CONTROLTYPE") + "," + GetDBValue(rowS, "FONTCOLOUR") 
					+ ", " + GetDBValue(rowS, "CAPTION") + ", " + GetDBValue(rowS, "FONTNAME") + ", " + GetDBValue(rowS, "FONTBOLD") 
					+ ", " + GetDBValue(rowS, "FONTITALIC") + ", " + GetDBValue(rowS, "FONTSIZE") + ", 0"
					+ ", " + GetDBValue(rowS, "SKIPCONDITION") + ", " + GetDBValue(rowS, "HEIGHT") + ", " + GetDBValue(rowS, "WIDTH") 
					+ ", " + GetDBValue(rowS, "CAPTIONX")	+ ", " + GetDBValue(rowS, "CAPTIONY") + ", " + GetDBValue(rowS, "X") 
					+ ", " + GetDBValue(rowS, "Y") + ", " + GetDBValue(rowS, "PRINTORDER") + ", " + GetDBValue(rowS, "HIDDEN") 
					+ ", " + GetDBValue(rowS, "LOCALFLAG") + ", " + GetDBValue(rowS, "OPTIONAL") + ", " + GetDBValue(rowS, "MANDATORY") 
					+ ", " + GetDBValue(rowS, "REQUIRECOMMENT") + ", " + GetDBValue(rowS, "ROLECODE") + ", 0, " + qGroupIdD 
					+ ", 0," + GetDBValue(rowS, "SHOWSTATUSFLAG") + ", " + GetDBValue(rowS, "CAPTIONFONTNAME") + ", " + GetDBValue(rowS, "CAPTIONFONTBOLD") + ", "
					+ GetDBValue(rowS, "CAPTIONFONTITALIC") + ", " + GetDBValue(rowS, "CAPTIONFONTSIZE") + ", " + GetDBValue(rowS, "CAPTIONFONTCOLOUR") 
					+ ", " + GetDBValue(rowS, "ELEMENTUSE") + ", " + GetDBValue(rowS, "DISPLAYLENGTH") + ", " + GetDBValue(rowS, "HOTLINK") + ")";
				
				//logging
				log.Info( "Inserting element:"
					+ " study:" + studyIdS + "->" + studyIdD
					+ " eform:" + rowS["CRFPAGEID"].ToString() + "->" + eFormIdD 
					+ " element:" + rowS["CRFELEMENTID"].ToString() + "->" + elementId );
				log.Debug( sql );

				RunSQL( dbConn, sql );

				if( qGroupIdD != "" )
				{
					//element is a question group - insert the qgroup row
					InsertEformQGroup( dbConn, studyIdS, studyIdD, rowS, eFormIdD, qGroupIdD  );
				}

				if( ( dataItemId != "0" ) || ( qGroupIdD != "0" ) )
				{
					//if we are adding an enterable element, renumber eform elements
					RenumberEformElements(dbConn, studyIdD, rowS["VERSIONID"].ToString(), eFormIdD );
				}

				//get the new row from the db
				ds = GetEformElements( dbConn, studyIdD, eFormIdD, elementId );
				if( ds.Tables[0].Rows.Count == 1 )
				{
					newRow = ds.Tables[0].Rows[0];
				}
		
				//logging
				log.Debug( "Database updated" );

				return( newRow );
			}
			finally
			{
				ds.Dispose();
			}
		}

		/// <summary>
		/// Copy question group (QGROUP table)
		/// </summary>
		/// <param name="dbConn"></param>
		/// <param name="studyIdD"></param>
		/// <param name="studyIdS"></param>
		/// <param name="rowS"></param>
		/// <param name="rowD"></param>
		public static void CopyQGroup( IDbConnection dbConn, string studyIdD, string studyIdS, DataRow rowS, DataRow rowD,
			ArrayList copyProperties )
		{
			DataSet ds = new DataSet();

			//logging
			log.Debug( "Updating database" );

			try
			{
				//get the question group properties
				ds = GetQGroup( dbConn, studyIdS, rowS );

				if( ds.Tables[0].Rows.Count == 1 )
				{
					DShowProgressEvent("Copying question group attributes, " + rowD["QGROUPID"], false, StudyCopyGlobal.LogPriority.Low);

					//get the question group row
					DataRow row = ds.Tables[0].Rows[0];

					//build the update sql
					string sql =
						"UPDATE QGROUP "
						+ "SET "
						+ GetSetPropertySql( row, copyProperties ) + " "
						+ "WHERE CLINICALTRIALID=" + studyIdD + " AND VERSIONID=" + rowD["VERSIONID"].ToString() + " "
						+ "AND QGROUPID=" + rowD["QGROUPID"].ToString();

					//logging
					log.Info( "Copying question group:"
						+ " study:" + studyIdS + "->" + studyIdD
						+ " eform:" + rowS["CRFPAGEID"].ToString() + "->" + rowD["CRFPAGEID"].ToString() 
						+ " question group:" + rowS["QGROUPID"].ToString() + "->" + rowD["QGROUPID"].ToString()
						+ " properties:" + GetCopyPropertyString( copyProperties ) );
					log.Debug( sql );

					//execute the update
					RunSQL( dbConn, sql );
			
					//logging
					log.Debug( "Database updated" );
				}
				else
				{
					//logging
					log.Info( "Unable to copy question group, not found in source study:"
						+ " study:" + studyIdS + "->" + studyIdD
						+ " question group:" + rowS["QGROUPID"].ToString() + "->" + rowD["QGROUPID"].ToString() );
				}
			}
			finally
			{
				ds.Dispose();
			}
		}

		/// <summary>
		/// Copy question group (EFORMQGROUP table)
		/// </summary>
		/// <param name="dbConn"></param>
		/// <param name="studyIdD"></param>
		/// <param name="studyIdS"></param>
		/// <param name="rowS"></param>
		/// <param name="rowD"></param>
		public static void CopyEformQGroup( IDbConnection dbConn, string studyIdD, string studyIdS, DataRow rowS, DataRow rowD,
			ArrayList copyProperties )
		{
			DataSet ds = new DataSet();

			//logging
			log.Debug( "Updating database" );

			try
			{
				//get the eform question group properties
				ds = GetEformQGroup( dbConn, studyIdS, rowS );

				if( ds.Tables[0].Rows.Count == 1 )
				{
					DShowProgressEvent("Copying eForm question group attributes, " + rowD["QGROUPID"], false, StudyCopyGlobal.LogPriority.Low);

					//get the question group row
					DataRow row = ds.Tables[0].Rows[0];

					//build the update sql
					string sql =
						"UPDATE EFORMQGROUP "
						+ "SET "
						+ GetSetPropertySql( row, copyProperties ) + " "
						+ "WHERE CLINICALTRIALID=" + studyIdD + " AND VERSIONID=" + rowD["VERSIONID"].ToString() + " "
						+ "AND CRFPAGEID=" + rowD["CRFPAGEID"].ToString() + " AND QGROUPID=" + rowD["QGROUPID"].ToString();

					//logging
					log.Info( "Copying eform question group:"
						+ " study:" + studyIdS + "->" + studyIdD
						+ " eform:" + rowS["CRFPAGEID"].ToString() + "->" + rowD["CRFPAGEID"].ToString() 
						+ " question group:" + rowS["QGROUPID"].ToString() + "->" + rowD["QGROUPID"].ToString() 
						+ " properties:" + GetCopyPropertyString( copyProperties ) );
					log.Debug( sql );

					//execute the update
					RunSQL( dbConn, sql );
			
					//logging
					log.Debug( "Database updated" );
				}
				else
				{
					//logging
					log.Info( "Unable to copy eform question group, not found in source study:"
						+ " study:" + studyIdS + "->" + studyIdD
						+ " eform:" + rowS["CRFPAGEID"].ToString() + "->" + rowD["CRFPAGEID"].ToString() 
						+ " question group:" + rowS["QGROUPID"].ToString() + "->" + rowD["QGROUPID"].ToString() );
				}
			}
			finally
			{
				ds.Dispose();
			}
		}

		/// <summary>
		/// Get a string of properties to copy from an arraylist
		/// </summary>
		/// <param name="al"></param>
		/// <returns></returns>
		private static string GetCopyPropertyString( ArrayList al )
		{
			string s = "";

			for( int n = 0; n < al.Count; n ++ )
			{
				s += al[n] + ", ";
			}

			return( ( s.EndsWith( ", " ) ) ? s.Substring( 0, ( s.Length - 2 ) ) : s );
		}

		public static void InsertQGroup( IDbConnection dbConn, string studyIdD, string studyIdS, DataRow rowS, out string newQGroupId )
		{
			DataSet ds = new DataSet();

			//logging
			log.Debug( "Updating database" );

			try
			{
				newQGroupId = "";

				//get the question group properties
				ds = GetQGroup( dbConn, studyIdS, rowS );

				if( ds.Tables[0].Rows.Count == 1 )
				{
					DShowProgressEvent("Inserting question group attributes, " + newQGroupId, false, StudyCopyGlobal.LogPriority.Low);

					//get the question group row
					DataRow row = ds.Tables[0].Rows[0];

					//get next unused qgroupid
					newQGroupId = GetNextQGroupId( dbConn, studyIdD );

					//build the update sql
					string sql =
						"INSERT INTO QGROUP "
						+ "(CLINICALTRIALID, VERSIONID, QGROUPID, QGROUPCODE, QGROUPNAME, DISPLAYTYPE) "
						+ "VALUES (" + studyIdD + ", " + GetDBValue(row, "VERSIONID") + ", " + newQGroupId + ", "
						+ GetDBValue(row, "QGROUPCODE" ) + ", " + GetDBValue(row, "QGROUPNAME" ) + ", "
						+ GetDBValue(row, "DISPLAYTYPE" ) + ")";

					//logging
					log.Info( "Inserting question group:"
						+ " study:" + studyIdS + "->" + studyIdD
						+ " question group:" + newQGroupId );
					log.Debug( sql );

					//execute the update
					RunSQL( dbConn, sql );
				
					//logging
					log.Debug( "Database updated" );
				}
				else
				{
					//logging
					log.Info( "Unable to insert question group, not found in source study:"
						+ " study:" + studyIdS + "->" + studyIdD
						+ " question group:" + rowS["QGROUPID"].ToString() + "->" + newQGroupId );
				}
			}
			finally
			{
				ds.Dispose();
			}
		}

		private static void InsertEformQGroup( IDbConnection dbConn, string studyIdS, string studyIdD, DataRow rowS, string eformIdD,
			string newQGroupId )
		{
			DataSet ds = new DataSet();
			
			//logging
			log.Debug( "Updating database" );

			try
			{
				//get the question group properties
				ds = GetEformQGroup( dbConn, studyIdS, rowS );

				if( ds.Tables[0].Rows.Count == 1 )
				{
					DShowProgressEvent("Inserting eForm question group attributes, " + newQGroupId, false, StudyCopyGlobal.LogPriority.Low);

					//get the question group row
					DataRow row = ds.Tables[0].Rows[0];

					//build the update sql
					string sql =
						"INSERT INTO EFORMQGROUP "
						+ "(CLINICALTRIALID, VERSIONID, CRFPAGEID, QGROUPID, BORDER, DISPLAYROWS, INITIALROWS, MINREPEATS, MAXREPEATS) "
						+ "VALUES (" + studyIdD + ", " + GetDBValue(row, "VERSIONID") + ", " + eformIdD + ", "
						+ newQGroupId + ", " + GetDBValue(row, "BORDER") + ", " + GetDBValue(row, "DISPLAYROWS") + ", "
						+ GetDBValue(row, "INITIALROWS") + ", " + GetDBValue(row, "MINREPEATS") + ", " + GetDBValue(row, "MAXREPEATS") + ")";

					//logging
					log.Info( "Inserting question group:"
						+ " study:" + studyIdS + "->" + studyIdD
						+ " eform:" + rowS["CRFPAGEID"].ToString() + "->" + eformIdD 
						+ " question group:" + rowS["QGROUPID"].ToString() + "->" + newQGroupId );
					log.Debug( sql );

					//execute the update
					RunSQL( dbConn, sql );
				
					//logging
					log.Debug( "Database updated" );
				}
				else
				{
					//logging
					log.Info( "Unable to insert question group, not found in source study:"
						+ " study:" + studyIdS + "->" + studyIdD
						+ " eform:" + rowS["CRFPAGEID"].ToString() + "->" + eformIdD 
						+ " question group:" + rowS["QGROUPID"].ToString() );
				}
			}
			finally
			{
				ds.Dispose();
			}
		}

		/// <summary>
		/// Insert category values for new dataitem
		/// </summary>
		/// <param name="dbConn"></param>
		/// <param name="studyIdS"></param>
		/// <param name="studyIdD"></param>
		/// <param name="rowS"></param>
		/// <param name="dataItemId"></param>
		/// <returns></returns>
		public static void InsertCategoryValues( IDbConnection dbConn, string studyIdS, string studyIdD, DataRow rowS, string dataItemId )
		{
			//logging
			log.Debug( "Updating database" );

			DShowProgressEvent("Inserting category values " + rowS["DATAITEMCODE"].ToString(), false, StudyCopyGlobal.LogPriority.Low);
	
			//set the 'update arezzo' flag
			string sql = "UPDATE STUDYDEFINITION SET AREZZOUPDATESTATUS = 1 WHERE CLINICALTRIALID = " + studyIdD;
			RunSQL( dbConn, sql );

			//get category values for this dataitem
			DataSet ds = GetCategoryValues( dbConn, studyIdS, rowS );

			foreach( DataRow row in ds.Tables[0].Rows )
			{
				//insert the new dataitem
				sql = 
					"INSERT INTO VALUEDATA "
					+ "(CLINICALTRIALID, VERSIONID, DATAITEMID, VALUEID, VALUECODE, ITEMVALUE, ACTIVE, VALUEORDER) "
					+ "VALUES (" + studyIdD + ", " + GetDBValue( rowS, "VERSIONID" ) + ", " + dataItemId + ", "
					+ GetDBValue( row, "VALUEID" ) + ", " + GetDBValue( row, "VALUECODE" ) + ", " + GetDBValue( row, "ITEMVALUE" ) + ", "
					+ GetDBValue( row, "ACTIVE" ) + ", " + GetDBValue( row, "VALUEORDER" ) + ")";
			
				//logging
				log.Info( "Inserting category value:"
					+ " study:" + studyIdS + "->" + studyIdD
					+ " dataitem:" + dataItemId 
					+ " valueid:" + row["VALUEID"].ToString() );
				log.Debug( sql );

				RunSQL( dbConn, sql );
			}

			//logging
			log.Debug( "Database updated" );
		}

		/// <summary>
		/// Insert the DataItem of a MACRO formatted study into a MACRO unformatted study
		/// </summary>
		/// <param name="dbConn"></param>
		/// <param name="studyIdS"></param>
		/// <param name="studyIdD"></param>
		/// <param name="rowS"></param>
		/// <param name="insertValidation"></param>
		/// <returns></returns>
		/// 28/04/2006	issue 2729 enter description when adding dataitem
		public static string InsertDataItem( IDbConnection dbConn, string studyIdS, string studyIdD, DataRow rowS, bool insertValidation, 
			string studyNameS )
		{
			//logging
			log.Debug( "Updating database" );

			DShowProgressEvent("Inserting DataItem " + rowS["DATAITEMCODE"].ToString(), false, StudyCopyGlobal.LogPriority.Low);
	
			//set the 'update arezzo' flag
			string sql = "UPDATE STUDYDEFINITION SET AREZZOUPDATESTATUS = 1 WHERE CLINICALTRIALID = " + studyIdD;
			RunSQL( dbConn, sql );

			//get next available dataitemid
			string newDataItemId = GetNextDataItemId( dbConn, studyIdD );

			//build the metaDescription field
			string metaDescription = GetDBValue( rowS, "DESCRIPTION" );
			if( metaDescription != "NULL" )
			{
				metaDescription = metaDescription.Substring( 0, ( metaDescription.Length - 1 ) ) + " ";
			}
			else
			{
				metaDescription = "'";
			}
			metaDescription += DateTime.Now.ToString() + " Dataitem copied from study " + studyNameS + "'";

			//insert the new dataitem
			sql = 
				"INSERT INTO DATAITEM "
				+ "(CLINICALTRIALID, VERSIONID, DATAITEMID, DATAITEMCODE, DATAITEMNAME, DATATYPE, DATAITEMFORMAT, UNITOFMEASUREMENT, "
				+ "DATAITEMLENGTH, DERIVATION, DATAITEMHELPTEXT, COPIEDFROMCLINICALTRIALID, COPIEDFROMVERSIONID, COPIEDFROMDATAITEMID, "
				+ "EXPORTNAME, DATAITEMCASE, CLINICALTESTCODE, MACROONLY, DESCRIPTION)"
				+ "VALUES (" + studyIdD + ", " + GetDBValue( rowS, "VERSIONID" ) + ", " + newDataItemId + ", "
				+ GetDBValue( rowS, "DATAITEMCODE" ) + ", " + GetDBValue( rowS, "DATAITEMNAME" ) + ", " + GetDBValue( rowS, "DATATYPE" ) + ", "
				+ GetDBValue( rowS, "DATAITEMFORMAT" ) + ", " + GetDBValue( rowS, "UNITOFMEASUREMENT" ) + ", "
				+ GetDBValue( rowS, "DATAITEMLENGTH" ) + ", " + GetDBValue( rowS, "DERIVATION" ) + ", " + GetDBValue( rowS, "DATAITEMHELPTEXT" ) + ", "
				+ "null, null, null, " + GetDBValue( rowS, "EXPORTNAME" ) + ", " + GetDBValue( rowS, "DATAITEMCASE" ) + ", "
				+ GetDBValue( rowS, "CLINICALTESTCODE" ) + ", " + GetDBValue( rowS, "MACROONLY" ) + ", " + metaDescription + ")";
			
			//logging
			log.Info( "Inserting dataitem:"
				+ " study:" + studyIdS + "->" + studyIdD
				+ " dataitem:" + newDataItemId );
			log.Debug( sql );

			RunSQL( dbConn, sql );

			if ( insertValidation )
			{
				//insert dataitem validation
				InsertDataItemValidation( dbConn, rowS, studyIdS, studyIdD, newDataItemId );
			}

			//logging
			log.Debug( "Database updated" );

			return( newDataItemId );
		}

		/// <summary>
		/// Does the dataitemcode exist in the passed study
		/// </summary>
		/// <param name="dbConn"></param>
		/// <param name="studyIdD"></param>
		/// <param name="dataItemCode"></param>
		/// <returns></returns>
		public static bool DataItemCodeExists( IDbConnection dbConn, string studyIdD, string dataItemCode )
		{
			DataSet ds = new DataSet();

			try
			{
				string sql = "SELECT COUNT(*) AS MATCHES "
					+ "FROM DATAITEM "
					+ "WHERE CLINICALTRIALID = " + studyIdD + " "
					+ "AND DATAITEMCODE = '" + dataItemCode + "'";

				//execute the query
				ds = GetDataSet( dbConn, sql );

				//return whether the dataitemcode was found or not 
				return( ( System.Convert.ToInt32( ( ds.Tables[0].Rows[0]["MATCHES"].ToString() ) ) > 0 ) );
			}
			finally
			{
				ds.Dispose();
			}
		}

		public static bool EformQGroupExists( IDbConnection dbConn, string studyIdD, string eformId, string qGroupId )
		{
			DataSet ds = new DataSet();

			try
			{
				string sql = "SELECT COUNT(*) AS MATCHES "
					+ "FROM EFORMQGROUP "
					+ "WHERE CLINICALTRIALID = " + studyIdD + " "
					+ "AND CRFPAGEID = " + eformId + " "
					+ "AND QGROUPID = " + qGroupId;

				//execute the query
				ds = GetDataSet( dbConn, sql );

				//return whether the dataitemcode was found or not 
				return( ( System.Convert.ToInt32( ( ds.Tables[0].Rows[0]["MATCHES"].ToString() ) ) > 0 ) );
			}
			finally
			{
				ds.Dispose();
			}
		}

		public static bool QGroupCodeExists( IDbConnection dbConn, string studyIdD, string studyIdS, string qGroupIdS, 
			out string qGroupCode, out string qGroupIdD )
		{
			DataSet ds = new DataSet();
			bool exists = false;

			try
			{
				qGroupCode = "";
				qGroupIdD = "";

				string sql = "SELECT QGROUPCODE "
					+ "FROM QGROUP "
					+ "WHERE CLINICALTRIALID = " + studyIdS + " "
					+ "AND QGROUPID = " + qGroupIdS;

				//execute the query
				ds = GetDataSet( dbConn, sql );

				if( ds.Tables[0].Rows.Count == 1 )
				{
					qGroupCode = ds.Tables[0].Rows[0]["QGROUPCODE"].ToString();

					sql = "SELECT COUNT(*) AS MATCHES "
						+ "FROM QGROUP "
						+ "WHERE CLINICALTRIALID = " + studyIdD + " "
						+ "AND QGROUPCODE = '" + qGroupCode + "'";

					//execute the query
					ds = GetDataSet( dbConn, sql );

					exists = ( System.Convert.ToInt32( ( ds.Tables[0].Rows[0]["MATCHES"].ToString() ) ) > 0 );

					if( exists )
					{
						sql = "SELECT QGROUPID "
							+ "FROM QGROUP "
							+ "WHERE CLINICALTRIALID = " + studyIdD + " "
							+ "AND QGROUPCODE = '" + qGroupCode + "'";

						//execute the query
						ds = GetDataSet( dbConn, sql );

						if( ds.Tables[0].Rows.Count == 1 )
						{
							qGroupIdD = ds.Tables[0].Rows[0]["QGROUPID"].ToString();
						}
					}
				}

				//return whether the dataitemcode was found or not 
				return( exists );
			}
			finally
			{
				ds.Dispose();
			}
		}

		private static string GetNextQGroupId( IDbConnection dbConn, string studyIdD )
		{
			DataSet ds = new DataSet();
			int nextID = System.Convert.ToInt32( _MINQGROUPID );

			try
			{
				string sql = "SELECT MAX(QGROUPID) AS MAXID "
					+ "FROM QGROUP "
					+ "WHERE CLINICALTRIALID = " + studyIdD + " "
					+ "AND QGROUPID >= " + _MINQGROUPID;

				//execute the query
				ds = GetDataSet( dbConn, sql );

				//return the next available id
				if (!ds.Tables[0].Rows[0].IsNull("MAXID"))
				{
					nextID = System.Convert.ToInt32( ds.Tables[0].Rows[0]["MAXID"] ) + 1;
				}
				return( nextID.ToString() );
			}
			finally
			{
				ds.Dispose();
			}
		}

		/// <summary>
		/// Get next dataitem id - the max id from dataitem, crfpage and studyvisit
		/// </summary>
		/// <param name="dbConn"></param>
		/// <param name="studyIdD"></param>
		/// <returns></returns>
		/// 03/05/2006 bug 2730 check studyvisit, crfpage and dataitem for next available dataitemid
		private static string GetNextDataItemId( IDbConnection dbConn, string studyIdD )
		{
			DataSet ds = new DataSet();
			int nextID = System.Convert.ToInt32( _MINDATAITEMID );

			try
			{
				//first check the dataitemid maximum value
				string sql = "SELECT MAX(DATAITEMID) AS MAXID "
					+ "FROM DATAITEM "
					+ "WHERE CLINICALTRIALID = " + studyIdD + " "
					+ "AND DATAITEMID >= " + nextID;

				//execute the query
				ds = GetDataSet( dbConn, sql );

				//return the next available id
				if (!ds.Tables[0].Rows[0].IsNull("MAXID"))
				{
					nextID = System.Convert.ToInt32( ds.Tables[0].Rows[0]["MAXID"] ) + 1;
				}

				//next check the crfpageid maximum value
				sql = "SELECT MAX(CRFPAGEID) AS MAXID "
					+ "FROM CRFPAGE "
					+ "WHERE CLINICALTRIALID = " + studyIdD + " "
					+ "AND CRFPAGEID >= " + nextID;

				//execute the query
				ds = GetDataSet( dbConn, sql );

				//return the next available id
				if( !ds.Tables[0].Rows[0].IsNull("MAXID") )
				{
					nextID = System.Convert.ToInt32( ds.Tables[0].Rows[0]["MAXID"] ) + 1;
				}
		
				//finally check the visitid maximum value
				sql = "SELECT MAX(VISITID) AS MAXID "
					+ "FROM STUDYVISIT "
					+ "WHERE CLINICALTRIALID = " + studyIdD + " "
					+ "AND VISITID >= " + nextID;

				//execute the query
				ds = GetDataSet( dbConn, sql );

				//return the next available id
				if (!ds.Tables[0].Rows[0].IsNull("MAXID"))
				{
					nextID = System.Convert.ToInt32( ds.Tables[0].Rows[0]["MAXID"] ) + 1;
				}

				return( nextID.ToString() );
			}
			finally
			{
				ds.Dispose();
			}
		}

		/// <summary>
		/// Copy the DataItem of a MACRO formatted study into a MACRO unformatted study
		/// </summary>
		/// <param name="dbConn"></param>
		/// <param name="studyIdS"></param>
		/// <param name="studyIdD"></param>
		/// <param name="rowS"></param>
		/// <param name="rowD"></param>
		/// <param name="insertValidation"></param>
		public static void CopyDataItem( IDbConnection dbConn, string studyIdS, string studyIdD, DataRow rowS, DataRow rowD,
			ArrayList copyProperties, bool insertValidation )
		{
			//logging
			log.Debug( "Updating database" );

			DShowProgressEvent("Copying dataItem attributes, " + rowD["DATAITEMID"] + " : " + rowD["DATAITEMNAME"], false, StudyCopyGlobal.LogPriority.Low);

			//set the 'update arezzo' flag
			string sql = "UPDATE STUDYDEFINITION SET AREZZOUPDATESTATUS = 1 WHERE CLINICALTRIALID = " + studyIdD;
			RunSQL( dbConn, sql );

			if ( copyProperties.Count > 0)
			{
				//build the the update sql
				sql = 
					"UPDATE DATAITEM "
					+ "SET "
					+ GetSetPropertySql( rowS, copyProperties ) + " "
					+ "WHERE CLINICALTRIALID=" + studyIdD + " AND DATAITEMID='" + rowD["DATAITEMID"].ToString() + "'";

				//logging
				log.Info( "Copying dataitem:"
					+ " study:" + studyIdS + "->" + studyIdD
					+ " dataitem:" + rowS["DATAITEMID"].ToString() + "->" + rowD["DATAITEMID"].ToString() 
					+ " properties:" + GetCopyPropertyString( copyProperties ) );
				log.Debug( sql );

				//execute the update
				RunSQL( dbConn, sql );
			}

			if ( insertValidation )
			{
				//insert dataitem validation
				InsertDataItemValidation( dbConn, rowS, studyIdS, studyIdD, rowD["DATAITEMID"].ToString() );
			}

			//logging
			log.Debug( "Database updated" );
		}

		/// <summary>
		/// Insert DataItem validation
		/// </summary>
		/// <param name="dbConn"></param>
		/// <param name="rowS"></param>
		/// <param name="studyIdS"></param>
		/// <param name="studyIdD"></param>
		/// <param name="dataItemIdD"></param>
		private static void InsertDataItemValidation( IDbConnection dbConn, DataRow rowS, string studyIdS, string studyIdD,  
			string dataItemIdD )
		{
			DataSet validationS = new DataSet();
			DataSet validationD = new DataSet();
			string sql = "";

			//logging
			log.Debug( "Updating database" );

			try
			{
				//get all the validation for this dataitem from the source and destination study
				validationS = GetDataItemValidation( dbConn, studyIdS, rowS["DATAITEMID"].ToString() );
				validationD = GetDataItemValidation( dbConn, studyIdD, dataItemIdD );

				//loop through each validation
				foreach( DataRow validationRowS in validationS.Tables[0].Rows )
				{
					bool exists = false;

					//first see if this validation already exists on the dataitem
					foreach ( DataRow validationRowD in validationD.Tables[0].Rows )
					{
						if ((validationRowS["DATAITEMVALIDATION"].ToString() == validationRowD["DATAITEMVALIDATION"].ToString())
							&& (validationRowS["VALIDATIONMESSAGE"].ToString() == validationRowD["VALIDATIONMESSAGE"].ToString()))
						{
							exists = true;
						}
					}

					if (!exists)
					{
						DShowProgressEvent("Inserting dataItem validation, " + validationRowS["DATAITEMVALIDATION"], false, StudyCopyGlobal.LogPriority.Low);

						//insert the validation
						sql = "INSERT INTO DATAITEMVALIDATION "
							+ "(CLINICALTRIALID, VERSIONID, DATAITEMID, VALIDATIONID, VALIDATIONTYPEID, DATAITEMVALIDATION, "
							+ "VALIDATIONMESSAGE) "
							+ "VALUES (" + studyIdD + ", " + GetDBValue(validationRowS, "VERSIONID") + ", " + dataItemIdD + ", "
							+ GetNextValidationID( dbConn, studyIdD, validationRowS["VERSIONID"].ToString(), dataItemIdD ) + ", " 
							+ GetDBValue(validationRowS, "VALIDATIONTYPEID") + ", " + GetDBValue(validationRowS, "DATAITEMVALIDATION") + ", " 
							+ GetDBValue(validationRowS, "VALIDATIONMESSAGE") + ")";

						//logging
						log.Info( "Inserting dataitem validation:"
							+ " study:" + studyIdS + "->" + studyIdD
							+ " dataitem:" + rowS["DATAITEMID"].ToString() );
						log.Debug( sql );

						RunSQL( dbConn, sql );
					}
				}

				//logging
				log.Debug( "Database updated" );
			}
			finally
			{
				validationS.Dispose();
				validationD.Dispose();
			}
		}

		private static string GetNextValidationID( IDbConnection dbConn, string studyID, string versionID, string dID)
		{
			DataSet ds = new DataSet();
			int nextID = 1;

			try
			{
				//build the query sql
				string sql = "SELECT MAX(VALIDATIONID) AS MAXID "
					+ "FROM DATAITEMVALIDATION "
					+ "WHERE CLINICALTRIALID = " + studyID + " AND VERSIONID = " + versionID + " "
					+ "AND DATAITEMID = " + dID;
				//execute the query
				ds = GetDataSet( dbConn, sql );
				//return the next available id
				if (!ds.Tables[0].Rows[0].IsNull("MAXID"))
				{
					nextID = System.Convert.ToInt32( ds.Tables[0].Rows[0]["MAXID"] ) + 1;
				}
				return( nextID.ToString() );
			}
			finally
			{
				ds.Dispose();
			}

		}

		/// <summary>
		/// Get the next available CRFElementID from the database
		/// </summary>
		/// <param name="studyID"></param>
		/// <param name="versionID"></param>
		/// <param name="crfPageID"></param>
		/// <returns></returns>
		private static string GetNextEformElementID( IDbConnection dbConn, string studyID, string versionID, string crfPageID)
		{
			DataSet ds = new DataSet();
			int nextID = 1;

			try
			{
				//build the query sql
				string sql = "SELECT MAX(CRFELEMENTID) AS MAXID "
					+ "FROM CRFELEMENT "
					+ "WHERE CLINICALTRIALID = " + studyID + " AND VERSIONID = " + versionID + " "
					+ "AND CRFPAGEID = " + crfPageID;
				//execute the query
				ds = GetDataSet( dbConn, sql );
				//return the next available id
				if (!ds.Tables[0].Rows[0].IsNull("MAXID"))
				{
					nextID = System.Convert.ToInt32(ds.Tables[0].Rows[0][0]) + 1;
				}
				return( nextID.ToString() );
			}
			finally
			{
				ds.Dispose();
			}

		}

		/// <summary>
		/// Returns visit/eforms dataset
		/// </summary>
		/// <param name="dbConn"></param>
		/// <param name="studyId"></param>
		/// <returns></returns>
		public static DataSet GetVisitEforms(IDbConnection dbConn, string studyId)
		{
			DShowProgressEvent("Getting visit/eform data from database", false, StudyCopyGlobal.LogPriority.Low);
			string sql = "";

			//build the query sql
			if( IMEDDataAccess.CalculateConnectionType( dbConn ) == IMEDDataAccess.ConnectionType.Oracle )
			{
				sql = "SELECT v.VISITNAME, v.VISITCODE, x.VISITID, x.CRFPAGEID "
					+ "FROM STUDYVISITCRFPAGE x, STUDYVISIT v "
					+ "WHERE x.CLINICALTRIALID = v.CLINICALTRIALID "
					+ "AND x.VISITID = v.VISITID "
					+ "AND x.CLINICALTRIALID = " + studyId + " "
					+ "ORDER BY VISITORDER";
			}
			else
			{
				sql = "SELECT v.VISITNAME, v.VISITCODE, x.VISITID, x.CRFPAGEID "
					+ "FROM STUDYVISITCRFPAGE x LEFT JOIN STUDYVISIT v "
					+ "ON x.CLINICALTRIALID = v.CLINICALTRIALID "
					+ "AND x.VISITID = v.VISITID "
					+ "WHERE x.CLINICALTRIALID = " + studyId + " "
					+ "ORDER BY VISITORDER";
			}

			//logging
			log.Debug( sql );

			//return the data set
			return (GetDataSet(dbConn, sql));
		}

		/// <summary>
		/// Return Eforms dataset
		/// </summary>
		/// <param name="dbConn"></param>
		/// <param name="studyId"></param>
		/// <returns></returns>
		public static DataSet GetEforms(IDbConnection dbConn, string studyId)
		{
			return( GetEforms( dbConn, studyId, "0" ) );
		}

		/// <summary>
		/// Return Eforms dataset
		/// </summary>
		/// <param name="dbConn"></param>
		/// <param name="studyId"></param>
		/// <param name="eformId"></param>
		/// <returns></returns>
		public static DataSet GetEforms(IDbConnection dbConn, string studyId, string eformId )
		{
			DShowProgressEvent("Getting eForm data from database", false, StudyCopyGlobal.LogPriority.Low);

			//build the query sql
			string sql = "SELECT CRFTITLE, CRFPAGEID, CRFPAGECODE, BACKGROUNDCOLOUR, CRFPAGELABEL, LOCALCRFPAGELABEL, DISPLAYNUMBERS, "
				+ "HIDEIFINACTIVE, EFORMWIDTH "
				+ "FROM CRFPAGE "
				+ "WHERE CLINICALTRIALID = " + studyId + " ";

			if( eformId != "0" )
			{
				sql += "AND CRFPAGEID = " + eformId;
			}

			sql += "ORDER BY CRFPAGEORDER";

			//logging
			log.Debug( sql );

			//return the result set
			return (GetDataSet(dbConn, sql));
		}

		/// <summary>
		/// Get question group data
		/// </summary>
		/// <param name="dbConn"></param>
		/// <param name="studyIdS"></param>
		/// <param name="rowS"></param>
		/// <returns></returns>
		private static DataSet GetEformQGroup( IDbConnection dbConn, string studyIdS, DataRow rowS )
		{
			DShowProgressEvent("Getting eform question group data from database", false, StudyCopyGlobal.LogPriority.Low);

			//build the query sql
			string sql = "SELECT * "
				+ "FROM EFORMQGROUP "
				+ "WHERE CLINICALTRIALID = " + studyIdS + " "
				+ "AND VERSIONID = " + rowS["VERSIONID"].ToString() + " "
				+ "AND CRFPAGEID = " + rowS["CRFPAGEID"].ToString() + " "
				+ "AND QGROUPID = " + rowS["QGROUPID"].ToString();

			//logging
			log.Debug( sql );

			//return the result set
			return( GetDataSet( dbConn, sql ) );
		}

		private static DataSet GetQGroup( IDbConnection dbConn, string studyIdS, DataRow rowS )
		{
			DShowProgressEvent("Getting question group data from database", false, StudyCopyGlobal.LogPriority.Low);

			//build the query sql
			string sql = "SELECT * "
				+ "FROM QGROUP "
				+ "WHERE CLINICALTRIALID = " + studyIdS + " "
				+ "AND VERSIONID = " + rowS["VERSIONID"].ToString() + " "
				+ "AND QGROUPID = " + rowS["QGROUPID"].ToString();

			//logging
			log.Debug( sql );

			//return the result set
			return( GetDataSet( dbConn, sql ) );
		}

		private static DataSet GetQGroupQuestion( IDbConnection dbConn, string studyIdS, DataRow rowS )
		{
			DShowProgressEvent("Getting question group questions data from database", false, StudyCopyGlobal.LogPriority.Low);

			//build the query sql
			string sql = "SELECT * "
				+ "FROM QGROUPQUESTION "
				+ "WHERE CLINICALTRIALID = " + studyIdS + " "
				+ "AND VERSIONID = " + rowS["VERSIONID"].ToString() + " "
				+ "AND QGROUPID = " + rowS["QGROUPID"].ToString();

			//logging
			log.Debug( sql );

			//return the result set
			return( GetDataSet( dbConn, sql ) );
		}

		/// <summary>
		/// Get category value data
		/// </summary>
		/// <param name="dbConn"></param>
		/// <param name="studyIdS"></param>
		/// <param name="rowS"></param>
		/// <returns></returns>
		public static DataSet GetCategoryValues( IDbConnection dbConn, string studyId, DataRow row )
		{
			DShowProgressEvent("Getting category value data from database", false, StudyCopyGlobal.LogPriority.Low);

			//build the query sql
			string sql = "SELECT * "
				+ "FROM VALUEDATA "
				+ "WHERE CLINICALTRIALID = " + studyId + " "
				+ "AND VERSIONID = " + row["VERSIONID"].ToString() + " "
				+ "AND DATAITEMID = " + row["DATAITEMID"].ToString();

			//logging
			log.Debug( sql );

			//return the result set
			return( GetDataSet( dbConn, sql ) );
		}

		/// <summary>
		/// Return DataItem validation dataset
		/// </summary>
		/// <param name="studyID"></param>
		/// <param name="dID"></param>
		/// <returns></returns>
		public static DataSet GetDataItemValidation( IDbConnection dbConn, string studyID, string dID )
		{
			DShowProgressEvent("Getting dataItem validation data from database", false, StudyCopyGlobal.LogPriority.Low);

			//build the query sql
			string sql = "SELECT CLINICALTRIALID, VERSIONID, DATAITEMID, VALIDATIONID, VALIDATIONTYPEID, DATAITEMVALIDATION, "
				+ "VALIDATIONMESSAGE "
				+ "FROM DATAITEMVALIDATION "
				+ "WHERE CLINICALTRIALID = " + studyID + " AND DATAITEMID = " + dID + " "
				+ "ORDER BY DATAITEMID";

			//logging
			log.Debug( sql );

			//return the result set
			return( GetDataSet( dbConn, sql ) );
		}

		/// <summary>
		/// Return CRF Elements dataset
		/// </summary>
		/// <param name="dbConn"></param>
		/// <param name="studyId"></param>
		/// <param name="eFormId"></param>
		/// <returns></returns>
		public static DataSet GetEformElements(IDbConnection dbConn, string studyId, string eFormId)
		{
			return( GetEformElements( dbConn, studyId, eFormId, "0" ) );
		}

		/// <summary>
		/// Return CRF Element dataset
		/// </summary>
		/// <param name="studyID"></param>
		/// <param name="crfPageID"></param>
		/// <returns></returns>
		public static DataSet GetEformElements(IDbConnection dbConn, string studyId, string eFormId, string elementId)
		{
			DShowProgressEvent("Getting eFormElement data from database", false, StudyCopyGlobal.LogPriority.Low);
			string sql = "";

			//build the query sql
			if( IMEDDataAccess.CalculateConnectionType( dbConn ) == IMEDDataAccess.ConnectionType.Oracle )
			{
				sql = "SELECT c.CRFPAGEID, d.DATAITEMNAME, c.CRFELEMENTID, c.VERSIONID, c.DATAITEMID, c.CONTROLTYPE, c.FONTCOLOUR, c.CAPTION, c.FONTNAME, c.FONTBOLD, "
					+ "c.FONTITALIC, c.FONTSIZE, c.FIELDORDER, c.SKIPCONDITION, c.HEIGHT, c.WIDTH, c.CAPTIONX, c.CAPTIONY, c.X, c.Y, "
					+ "c.PRINTORDER, c.HIDDEN, c.LOCALFLAG, c.OPTIONAL, c.MANDATORY, c.REQUIRECOMMENT, c.ROLECODE, c.OWNERQGROUPID, "
					+ "c.QGROUPID, c.QGROUPFIELDORDER, c.SHOWSTATUSFLAG, c.CAPTIONFONTNAME, c.CAPTIONFONTBOLD, c.CAPTIONFONTITALIC, "
					+ "c.CAPTIONFONTSIZE, c.CAPTIONFONTCOLOUR, c.ELEMENTUSE, c.DISPLAYLENGTH, c.HOTLINK, d.DATATYPE, "
					+ "d.DATAITEMFORMAT, d.UNITOFMEASUREMENT, d.DATAITEMLENGTH, d.DERIVATION, d.DATAITEMHELPTEXT, "
					+ "d.EXPORTNAME, d.DATAITEMCASE, d.DESCRIPTION, d.DATAITEMCODE, d.CLINICALTESTCODE, d.MACROONLY, e.QGROUPCODE "
					+ "FROM CRFELEMENT c, DATAITEM d, QGROUP e "
					+ "WHERE c.CLINICALTRIALID = d.CLINICALTRIALID (+) AND "
					+ "c.VERSIONID = d.VERSIONID (+) AND "
					+ "c.DATAITEMID = d.DATAITEMID (+) AND "
					+ "c.QGROUPid = e.QGROUPid (+) AND "
					+ "c.CLINICALTRIALID = e.CLINICALTRIALID (+) AND "
					+ "c.CLINICALTRIALID = " + studyId + " AND "
					+ "c.CRFPAGEID = " + eFormId + " ";
			}
			else
			{
				sql = "SELECT c.CRFPAGEID, d.DATAITEMNAME, c.CRFELEMENTID, c.VERSIONID, c.DATAITEMID, c.CONTROLTYPE, c.FONTCOLOUR, c.CAPTION, c.FONTNAME, c.FONTBOLD, "
					+ "c.FONTITALIC, c.FONTSIZE, c.FIELDORDER, c.SKIPCONDITION, c.HEIGHT, c.WIDTH, c.CAPTIONX, c.CAPTIONY, c.X, c.Y, "
					+ "c.PRINTORDER, c.HIDDEN, c.LOCALFLAG, c.OPTIONAL, c.MANDATORY, c.REQUIRECOMMENT, c.ROLECODE, c.OWNERQGROUPID, "
					+ "c.QGROUPID, c.QGROUPFIELDORDER, c.SHOWSTATUSFLAG, c.CAPTIONFONTNAME, c.CAPTIONFONTBOLD, c.CAPTIONFONTITALIC, "
					+ "c.CAPTIONFONTSIZE, c.CAPTIONFONTCOLOUR, c.ELEMENTUSE, c.DISPLAYLENGTH, c.HOTLINK, d.DATATYPE, "
					+ "d.DATAITEMFORMAT, d.UNITOFMEASUREMENT, d.DATAITEMLENGTH, d.DERIVATION, d.DATAITEMHELPTEXT, "
					+ "d.EXPORTNAME, d.DATAITEMCASE, d.DESCRIPTION, d.DATAITEMCODE, d.CLINICALTESTCODE, d.MACROONLY, e.QGROUPCODE "
					+ "FROM CRFELEMENT c LEFT OUTER JOIN DATAITEM d "
					+ "ON d.CLINICALTRIALID = c.CLINICALTRIALID "
					+ "AND d.VERSIONID = c.VERSIONID "
					+ "AND d.DATAITEMID = c.DATAITEMID "
					+ "LEFT OUTER JOIN QGROUP e "
					+ "ON c.QGROUPID = e.QGROUPID "
					+ "AND c.CLINICALTRIALID = e.CLINICALTRIALID "
					+ "WHERE c.CLINICALTRIALID = " + studyId + " AND c.CRFPAGEID = " + eFormId + " ";
			}

			if( elementId != "0" )
			{
				sql += "AND CRFELEMENTID = " + elementId + " ";
			}
				
			sql += "ORDER BY c.CRFELEMENTID";

			//logging
			log.Debug( sql );

			//return the data set
			return (GetDataSet(dbConn, sql));
		}

		public static DataSet GetDataItems(IDbConnection dbConn, string studyId, string dataType)
		{
			DShowProgressEvent("Getting dataItem data from database", false, StudyCopyGlobal.LogPriority.Low);

			string sql = "SELECT * FROM DATAITEM WHERE CLINICALTRIALID = " + studyId + " AND DATATYPE = " + dataType;

			//logging
			log.Debug( sql );

			return (GetDataSet(dbConn, sql));
		}

		/// <summary>
		/// Returns a dataset of MACRO databases
		/// </summary>
		/// <param name="connectionString">Database connection string</param>
		/// <returns>Dataset of results</returns>
		public static DataSet GetDatabaseList(string connectionString)
		{
			string sql = "SELECT DATABASECODE, DATABASETYPE, DATABASEUSER, DATABASEPASSWORD, SERVERNAME, "
				+ "NAMEOFDATABASE FROM DATABASES";

			//logging
			log.Debug( sql );

			return( GetDataSet( connectionString, sql ) );
		}

		/// <summary>
		/// Returns a dataset of MACRO studies
		/// </summary>
		/// <param name="connectionString">Database connection string</param>
		/// <returns>Dataset of results</returns>
		public static DataSet GetStudyList(string connectionString)
		{
			string sql = "SELECT CLINICALTRIALNAME, CLINICALTRIALID FROM CLINICALTRIAL WHERE CLINICALTRIALID > 0 "
				+ "ORDER BY CLINICALTRIALNAME";

			//logging
			log.Debug( sql );

			return( GetDataSet(connectionString, sql) );
		}

		private static string GetDBValue(DataRow row, string colName)
		{
			if (row.IsNull(colName))
			{
				return ("NULL");
			}
			else
			{
				string s = row[colName].ToString();
				if (row[colName].GetType().ToString() == "System.String") 
				{
					s = s.Replace("'", "''");
					s = s.Replace("\"", "\"\"");
					s = "'" + s + "'";
				}
				return (s);
			}
		}

		private static string ConvertToUnicode(string s)
		{
			UnicodeEncoding uni = new UnicodeEncoding();

			Byte[] encoded = uni.GetBytes(s);
			Byte[] con = UnicodeEncoding.Convert(UnicodeEncoding.UTF8, UnicodeEncoding.Unicode, encoded);
			return (uni.GetString(con));
		}

		private static string ConvertToAscii(string s)
		{
			ASCIIEncoding ascii = new ASCIIEncoding();
			Byte[] encodedBytes = ascii.GetBytes(s);
			return( ascii.GetString(encodedBytes));
		}

		private static void RunSQL(IDbConnection dbConn, string sql)
		{
			IMEDDataAccess imedData = new IMEDDataAccess();
			IDbCommand dbCommand = null;

			try
			{
				// get command
				dbCommand = imedData.GetCommand(dbConn);
				// create command text
				dbCommand.CommandText = sql;
				dbCommand.ExecuteNonQuery();
			}
			finally
			{
				//clean up objects
				dbCommand.Dispose();
			}
		}

		/// <summary>
		/// Returns a dataset from the database
		/// </summary>
		/// <param name="connectionString">MACRO database connection string</param>
		/// <param name="sql">SQL to run against the MACRO database</param>
		/// <returns>Dataset containing MACRO data</returns>
		private static DataSet GetDataSet(string connectionString, string sql)
		{
			IMEDDataAccess imedData = new IMEDDataAccess();
			IDbConnection dbConn = null;
			
			try
			{
				// open ImedDataAccess class
				IMEDDataAccess.ConnectionType imedConnType;
				// calculate connection type
				imedConnType=IMEDDataAccess.CalculateConnectionType( connectionString );
				// get connection & open
				dbConn = imedData.GetConnection( imedConnType, connectionString );
				dbConn.Open();
				
				return( GetDataSet( dbConn, sql ) );
			}
			finally
			{
				//clean up objects
				dbConn.Close();
				dbConn.Dispose();
			}
		}

		/// <summary>
		/// Returns a dataset from the database
		/// </summary>
		/// <param name="dbConn">Database connection object</param>
		/// <param name="sql">SQL to run against the MACRO database</param>
		/// <returns>Dataset containing MACRO data</returns>
		public static DataSet GetDataSet(IDbConnection dbConn, string sql)
		{
			IMEDDataAccess imedData = new IMEDDataAccess();
			IDbCommand dbCommand = null;
			IDbDataAdapter dbDataAdapter = null;
			DataSet dataDB = new DataSet();

			try
			{
				// get command
				dbCommand = imedData.GetCommand( dbConn );
				// create command text
				dbCommand.CommandText = sql;
				// get data adaptor
				dbDataAdapter = imedData.GetDataAdapter( dbCommand );
				// fill dataset
				dbDataAdapter.Fill( dataDB );
				
				return( dataDB );
			}
			finally
			{
				//clean up objects
				dbCommand.Dispose();
			}
		}
	}
}
