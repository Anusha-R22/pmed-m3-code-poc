<%@ Language=VBScript %>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<!-- METADATA TYPE="typelib" uuid="00000205-0000-0010-8000-00AA006D2EA4" --> 
<%
'*************************
' Header block
'*************************

sReportTitle = "Questions"

sIncludeVML = 0 'Don't include VML styles

%>
<!--#include file="report_initialise.asp" -->
<!--#include file="report_open_macro_database.asp" -->
<%

' RS 20-AUG-03: Set content type so that XML is displayed
Response.ContentType = "text/xml"

' DPH 01/07/2005 - 
dim clinTrialId
clinTrialId = Request.QueryString("clinicaltrialid")

'*************************
' Content block
'*************************

Set rsResult = CreateObject("ADODB.Recordset")
Set rsResult1 = CreateObject("ADODB.Recordset")
Connect.CursorLocation = adUseClient 

Set adoCmd = Server.CreateObject("ADODB.Command") 
Set adoCmd.ActiveConnection = Connect
Dim sQuery 

dim docXML
dim docXSL
dim docXMLOut

Dim root 
Dim newNode 

' DPH 22/01/2004 - Allow save of file to server directory
dim sXMLFile
dim oStream
dim sContents

set docXML = Server.CreateObject("MSXML2.DOMDocument.3.0")
set docXSL = Server.CreateObject("MSXML2.DOMDocument.3.0")
set docXMLOut = Server.CreateObject("MSXML2.DOMDocument.3.0")


docXML.ValidateonParse = True
docXSL.ValidateonParse = True
docXMLOut.ValidateonParse = True
docXMLOut.loadXML "<root created='" & date() & " " & time() & "' />"
set root = docXMLOut.documentElement


sQuery =  " Select *"
sQuery = sQuery & " from clinicaltrial "
sQuery = sQuery & " where clinicaltrialid = " & clinTrialId & " "

AddNode sQuery, "ClinicalTrial", docXML

' DPH 04/07/2005 - Only show visits with eForms
sQuery =  " Select distinct studyvisit.*"
sQuery = sQuery & " from studyvisit,studyvisitcrfpage "
sQuery = sQuery & " where studyvisit.clinicaltrialid = studyvisitcrfpage.clinicaltrialid"
sQuery = sQuery & " and studyvisit.visitid= studyvisitcrfpage.visitid"
sQuery = sQuery & " and studyvisit.clinicaltrialid = " & clinTrialId & " order by studyvisit.visitorder"
 
AddNode sQuery, "StudyVisit", docXML

sQuery =  "Select studyvisit.visitid,studyvisit.visitcode,studyvisit.visitorder,studyvisit.visitname,studyvisitcrfpage.*  "
sQuery = sQuery & "from  studyvisit,studyvisitcrfpage "
sQuery = sQuery & " where studyvisit.clinicaltrialid = studyvisitcrfpage.clinicaltrialid"
sQuery = sQuery & "   and studyvisit.visitid= studyvisitcrfpage.visitid "
sQuery = sQuery & "  and studyvisit.clinicaltrialid = " & clinTrialId & " order by visitorder"

AddNode sQuery, "StudyVisitCRFPage", docXML

' DPH 19/10/2005 - get groups - either only on eform or RQG
sQuery = "select * from " 
sQuery = sQuery & "(SELECT DISTINCT OwnerQGroupId,crfe.crfpageid,crf.crfpagecode,qgroupcode,qgroupname,crf.crftitle "
sQuery = sQuery & "FROM CRFELEMENT crfe, qgroup qg, crfpage crf "
sQuery = sQuery & "WHERE crfe.OwnerQGroupId <> 0 AND crfe.clinicaltrialid = " & clinTrialId & " "
sQuery = sQuery & "and crfe.clinicaltrialid = qg.clinicaltrialid and crfe.clinicaltrialid = crf.clinicaltrialid "
sQuery = sQuery & "and crfe.crfpageid = crf.crfpageid and crfe.ownerqgroupid = qg.qgroupid "
sQuery = sQuery & "and crfe.dataitemid > 0 "
sQuery = sQuery & "union "
sQuery = sQuery & "SELECT DISTINCT 0,crfe.crfpageid,crfpagecode,crfpagecode,crftitle,crftitle "
sQuery = sQuery & "FROM CRFELEMENT crfe,crfpage crf "
sQuery = sQuery & "WHERE crfe.OwnerQGroupId = 0 AND crfe.clinicaltrialid = " & clinTrialId & " "
sQuery = sQuery & "and crfe.clinicaltrialid = crf.clinicaltrialid "
sQuery = sQuery & "and crfe.crfpageid = crf.crfpageid and crfe.dataitemid > 0) d "
sQuery = sQuery & "order by d.crfpageid "

AddNode sQuery, "PageGroups", docXML

' DPH 04/07/2005 - Only retrieve eforms with question elements
sQuery = " Select distinct crfpage.* "
sQuery = sQuery & " from crfpage,crfelement,dataitem "
sQuery = sQuery & " where crfpage.clinicaltrialid = crfelement.clinicaltrialid"
sQuery = sQuery & "   and crfpage.crfpageid = crfelement.crfpageid"
sQuery = sQuery & "   and crfelement.clinicaltrialid = dataitem.clinicaltrialid "
sQuery = sQuery & "   and crfelement.dataitemid = dataitem.dataitemid"
' screen out non-questions
sQuery = sQuery & "   and crfelement.dataitemid > 0"
' screen out multimedia
sQuery = sQuery & "   and dataitem.datatype <> 5"
sQuery = sQuery & "   and crfpage.clinicaltrialid = " & clinTrialId & " "
 
AddNode sQuery, "CRFPage", docXML

sQuery = " Select crfpage.crfpageid,CRFPageCode,CRFTitle,CRFPageOrder,FieldOrder,Mandatory,crfelement.DataItemId,DataItemCode,OwnerQGroupId"
sQuery = sQuery & " from crfpage,crfelement,dataitem "
sQuery = sQuery & " where crfpage.clinicaltrialid = crfelement.clinicaltrialid"
sQuery = sQuery & "   and crfpage.crfpageid= crfelement.crfpageid "
sQuery = sQuery & "   and crfelement.clinicaltrialid = dataitem.clinicaltrialid "
sQuery = sQuery & "   and crfelement.dataitemid = dataitem.dataitemid"
' screen out multimedia
sQuery = sQuery & "   and dataitem.datatype <> 5"
sQuery = sQuery & "   and crfpage.clinicaltrialid = " & clinTrialId & " "

AddNode sQuery, "CRFElement", docXML

' DPH 18/10/2005 - get RQG detail
sQuery = "SELECT DISTINCT QGROUPQUESTION.QGROUPID, QGROUP.QGROUPCODE, QGROUPQUESTION.DATAITEMID, QGROUPQUESTION.QORDER, CRFELEMENT.MANDATORY, CRFELEMENT.CRFPAGEID, DATAITEM.DATAITEMCODE "
sQuery = sQuery & "FROM QGROUPQUESTION, QGROUP, CRFELEMENT, DATAITEM "
sQuery = sQuery & "WHERE QGROUPQUESTION.CLINICALTRIALID = " & clinTrialId & " AND (QGROUPQUESTION.CLINICALTRIALID = QGROUP.CLINICALTRIALID) "
sQuery = sQuery & "AND (QGROUPQUESTION.CLINICALTRIALID = CRFELEMENT.CLINICALTRIALID) "
sQuery = sQuery & "AND (QGROUPQUESTION.CLINICALTRIALID = DATAITEM.CLINICALTRIALID) AND (QGROUPQUESTION.DATAITEMID = DATAITEM.DATAITEMID) "
sQuery = sQuery & "AND (QGROUPQUESTION.DATAITEMID = CRFELEMENT.DATAITEMID) AND (QGROUPQUESTION.QGROUPID = QGROUP.QGROUPID) "
sQuery = sQuery & "ORDER BY QGROUPQUESTION.QGROUPID, CRFELEMENT.CRFPAGEID, QGROUPQUESTION.QORDER "

AddNode sQuery, "RQGDetail", docXML

sQuery = " Select *"
sQuery = sQuery & " from dataitem "
' screen out multimedia
sQuery = sQuery & " where dataitem.datatype <> 5"
sQuery = sQuery & " and clinicaltrialid = " & clinTrialId & " "

AddNode sQuery, "DataItem", docXML

sQuery =  " Select Distinct DataItem1.DataItemId, DataItem1.DataItemCode "
sQuery = sQuery & " from dataitem dataitem1,valuedata "
sQuery = sQuery & " where dataitem1.clinicaltrialid = valuedata.clinicaltrialid"
sQuery = sQuery & "   and dataitem1.dataitemid= valuedata.dataitemid "
sQuery = sQuery & "   and dataitem1.clinicaltrialid = " & clinTrialId & " "

AddNode sQuery, "CodeLists", docXML

sQuery =  " Select DataItem1.DataItemId, DataItem1.DataItemCode,ValueCode,ItemValue "
sQuery = sQuery & " from dataitem dataitem1,valuedata "
sQuery = sQuery & " where dataitem1.clinicaltrialid = valuedata.clinicaltrialid"
sQuery = sQuery & "   and dataitem1.dataitemid= valuedata.dataitemid "
sQuery = sQuery & "   and dataitem1.clinicaltrialid = " & clinTrialId & " "

AddNode sQuery, "ValueData", docXML

sQuery =  " Select trialsite,personid,localidentifier1 from trialsubject "
sQuery = sQuery & "   where trialsubject.clinicaltrialid = " & clinTrialId & " "

AddNode sQuery, "TrialSubject", docXML

sQuery =  " Select trialsite,personid,visitid,visittaskid,visitcyclenumber from visitinstance "
sQuery = sQuery & "   where visitinstance.clinicaltrialid = " & clinTrialId & " "

AddNode sQuery, "VisitInstance", docXML

sQuery =  " Select trialsite,personid,visitid,visitcyclenumber,crfpageid,crfpagecyclenumber,crfpagetaskid from CRFPageInstance "
sQuery = sQuery & "   where CRFPageInstance.clinicaltrialid = " & clinTrialId & " "

AddNode sQuery, "CRFPageInstance", docXML

sQuery =  " Select trialsite,personid,crfpagetaskid,dataitemid,responsevalue,repeatnumber from DataItemResponse "
sQuery = sQuery & "   where DataItemResponse.clinicaltrialid = " & clinTrialId & " "

AddNode sQuery, "DataItemResponse", docXML

' DPH 23/11/2005 - Get group repeat info for a particular page
sQuery = "SELECT * FROM ( "
sQuery = sQuery & "SELECT DISTINCT dir.PERSONID, dir.CRFPAGETASKID, qq.QGROUPID, dir.REPEATNUMBER "
sQuery = sQuery & "FROM DATAITEMRESPONSE dir, QGROUPQUESTION qq "
sQuery = sQuery & "WHERE dir.CLINICALTRIALID = " & clinTrialId & " "
sQuery = sQuery & "AND dir.DATAITEMID = qq.DATAITEMID AND dir.CLINICALTRIALID = qq.CLINICALTRIALID "
sQuery = sQuery & "UNION "
sQuery = sQuery & "SELECT DISTINCT dir.PERSONID, dir.CRFPAGETASKID, ce.OWNERQGROUPID, dir.REPEATNUMBER "
sQuery = sQuery & "FROM DATAITEMRESPONSE dir, CRFELEMENT ce "
sQuery = sQuery & "WHERE dir.CLINICALTRIALID = " & clinTrialId & " "
sQuery = sQuery & "AND dir.DATAITEMID = ce.DATAITEMID AND dir.CLINICALTRIALID = ce.CLINICALTRIALID AND ce.OWNERQGROUPID = 0 "
sQuery = sQuery & ") di ORDER BY di.PERSONID, di.CRFPAGETASKID, di.QGROUPID, di.REPEATNUMBER "

AddNode sQuery, "ResponseGroupInfo", docXML


docXSL.load server.mappath("cdisc_output.xsl")

docXML.LoadXML docXMLOut.transformnode( docXSL )
'docXML.save response

' dph - save file to server as xml
sXMLFile = server.mappath("cdisc_output.xsl")
sXMLFile = left(sXMLFile,len(sXMLFile)-3) & "xml"
docXML.save sXMLFile
'docXMLOut.save sXMLFile

set docXML = Nothing
set docXSL = Nothing
set docXMLOut = Nothing

Response.Redirect "cdisc_show.htm"

sub SQLToResponse( sQuery )

adoCmd.CommandText = sQuery
rsResult.open sQuery, connect
rsResult.Save Response, adPersistXML 


end sub

sub SQLToFile( sQuery, sFile )

adoCmd.CommandText = sQuery
rsResult.open sQuery, connect
rsResult.Save sFile, adPersistXML 

end sub

sub AddNode( sQuery, sName, xDOM )

adoCmd.CommandText = ucase(sQuery)
'response.write sQuery & "<BR>"
'response.Flush()
rsResult.open ucase(sQuery), connect
rsResult.Save docXML, adPersistXML
rsResult.Close

set newNode = xDOM.createNode(1, sName , "")
newNode.appendchild(docXML.documentElement)
root.appendChild(newNode)

end sub


'*************************
' Footer block
'*************************

%>
<!--#include file="report_close.asp" -->

