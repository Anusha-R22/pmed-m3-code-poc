<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Discrepancy Details"

if request.querystring("chart") > "" then
		sIncludeVML = 1 'Include VML styles
else
		sIncludeVML = 0 'Don't include VML styles
end if

%><!--#include file="report_initialise.asp" -->
<!--#include file="report_open_macro_database.asp" -->
<!--#include file="report_functions.asp" -->
<!--#include file="report_header.asp" -->
  <%


'*************************
' Parameter Check
'*************************
if request.QueryString("clinicaltrialname")="" or request.querystring("trialsite")="" or request.querystring("mimessagestatus")="" then
	WritePara "Missing Parameter, please check whether the following parameters are passed:"
	WritePara "clinicaltrialid, trialsite, mimessagestatus"
	response.end
else
	pTrialName = request.QueryString("clinicaltrialname")
	pTrialSite = request.querystring("trialsite")
	pMessageStatus = request.querystring("mimessagestatus")
	pOC = Request.QueryString("oc") ' "true", "false", or ""
end if
	

'*************************
' Content block
'*************************
sHeader = pTrialName
sHeader = sHeader & "&nbsp;&nbsp;&nbsp;Site: " & pTrialSite
sHeader = sHeader & "&nbsp;&nbsp;&nbsp;"
select case CInt(pMessageStatus)
	case 0:	sHeader = sHeader & "Discrepancy Status: Raised"
	case 1: sHeader = sHeader & "Discrepancy Status: Responded"
	case 2: sHeader = sHeader & "Discrepancy Status: Closed"
end select

WriteGroupHeader "Study",sHeader
Set rsResult = CreateObject("ADODB.Recordset")
Set rsResult1 = CreateObject("ADODB.Recordset")

'*************************
' Get all discrepancies of given status
'*************************
sQuery = " select mimessage.*, dataitem.dataitemname, crfpage.crftitle, trialsubject.localidentifier1"
sQuery = sQuery & " from mimessage, dataitem, crfpage, trialsubject"
sQuery = sQuery & " where dataitem.dataitemid = mimessagedataitemid and dataitem.clinicaltrialid = (select clinicaltrialid from clinicaltrial where clinicaltrialname = mimessagetrialname)"
sQuery = sQuery & " and crfpage.crfpageid = mimessagecrfpageid  and crfpage.clinicaltrialid = (select clinicaltrialid from clinicaltrial where clinicaltrialname = mimessagetrialname)"
' DPH 22/07/2004 - collect subject label detail and display in report
sQuery = sQuery & " and trialsubject.clinicaltrialid = (select clinicaltrialid from clinicaltrial where clinicaltrialname = mimessagetrialname) and trialsubject.trialsite = mimessage.mimessagesite"
sQuery = sQuery & " and trialsubject.personid = mimessage.mimessagepersonid"
sQuery = sQuery & " and mimessagehistory=0 "
sQuery = sQuery & " and mimessagetype = 0"
sQuery = sQuery & " and mimessagestatus = " & pMessageStatus
sQuery = sQuery & " and MIMessageTrialName = '" & pTrialName & "'"
sQuery = sQuery & " and MIMessageSite = '" & pTrialSite & "'"
if pOC = "true" then
	sQuery = sQuery & " and MIMessageOCDiscrepancyId <> 0"
elseif pOC = "false" then
	sQuery = sQuery & " and MIMessageOCDiscrepancyId = 0"
end if
sQuery = sQuery & " order by MIMessageTrialname, MIMessageSite, MIMessageType, MIMessageStatus"

'response.write sQuery
'response.flush

rsResult.Open sQuery,Connect,3

if pOC = "" then
	'determine whether the recordset includes both OC- and MACRO-originating discrepancies
	if clng(rsResult.Fields("MIMessageOCDiscrepancyId").Value) = 0 then
		sCriterion = "<>"
	else
		sCriterion = "="
	end if
	rsResult.Find "MIMessageOCDiscrepancyId " & sCriterion & " 0"
end if
if not (pOC = "" and rsResult.EOF) then
	Response.Write "<form id=oc method=get action='" & Request.ServerVariables("URL") & "'><table bgcolor=#eeeeee cellpadding=3><tr><td>" & HiddenFormElements(Array("oc"))
	Response.Write "<select name=oc><option value=''>All discrepancies</option><option value=true"
	if pOC = "true" then
		Response.Write " selected"
	end if
	Response.Write ">Discrepancies originating in OC</option><option value=false"
	if pOC = "false" then
		Response.Write " selected"
	end if
	Response.Write ">Discrepancies originating in MACRO</option></select></td><td>"
	Response.Write "<input type=submit value=Filter></td></tr></table></form>"
end if

if not rsResult.BOF then
	rsResult.MoveFirst
end if

WriteTableStart
WriteTableRowStart
WriteHeaderCell "Study"
WriteHeaderCell "Site"
WriteHeaderCell "Subject"
WriteHeaderCell "eForm"
WriteHeaderCell "Question"
WriteHeaderCell "Response"
WriteHeaderCell "Discrepancy Message"
WriteHeaderCell "Priority"
WriteHeaderCell "Status"
WriteTableRowEnd

do until rsResult.eof 
	WriteTableRowStart
	WriteCell pTrialname
	WriteCell pTrialsite
	WriteCell fIdOrLabel(rsResult("mimessagepersonid"), rsResult("localidentifier1"))
	WriteCell rsResult("crftitle")
	WriteCell rsResult("dataitemname")
	WriteCell rsResult("mimessageresponsevalue")
	WriteCell rsResult("mimessagetext")
	WriteCell rsResult("mimessagepriority")
	' dph 15/03/2004 - display discrepancy status
	select case CInt(pMessageStatus)
		case 0:	WriteCell "Raised"
		case 1: WriteCell "Responded"
		case 2: WriteCell "Closed"
	end select	
	'WriteCell pMessageStatus
	WriteTableRowEnd
	rsResult.movenext
loop


rsResult.Close
set RsResult = Nothing
set RsResult1 = Nothing


'*************************
' Footer block
'*************************

%>
  <!--#include file="report_footer.asp" -->
<!--#include file="report_close.asp" -->