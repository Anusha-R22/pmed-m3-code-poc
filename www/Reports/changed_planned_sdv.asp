<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Done SDVs with Changed Remote Site Data"

if request.querystring("chart") > "" then
		sIncludeVML = 1 'Include VML styles
else
		sIncludeVML = 0 'Don't include VML styles
end if

%>
<!--#include file="report_initialise.asp" -->
<!--#include file="report_open_macro_database.asp" -->
<!--#include file="report_functions.asp" -->
<!--#include file="report_header.asp" -->
  <%
dim lStudyId
dim sSite
dim lSubjectId
WriteSelectionHeader Array(), true, Array(), true
if Request.QueryString("study") <> "" then
	lStudyId = clng(Request.QueryString("study"))
else
	lStudyId = 0
end if
sSite = Request.QueryString("site")
if Request.QueryString("subject") <> "" then
	dim asSplit
	asSplit = split(Request.QueryString("subject"), "`")
	lStudy = clng(asSplit(0))
	sSite = asSplit(1)
	lSubjectId = clng(asSplit(2))
else
	lSubjectId = 0
end if

'******************************************
' RS 11AUG2003 Bail out of no studies exist
'******************************************
if sPermittedStudyList="" then
	WriteGroupHeader "Available Studies","You do not have access to any studies"
	%>
  <!--#include file="report_close.asp" -->
  <%
	Response.end()
end if

'*************************
' Content block
'*************************

Set rsResult = CreateObject("ADODB.Recordset")
Set rsResult1 = CreateObject("ADODB.Recordset")

sQuery = "Select distinct c.clinicaltrialid,clinicaltrialname from clinicaltrial c "
if lStudyId = 0 then
	sQuery = sQuery & " where c.clinicaltrialid in (" & sPermittedStudyList & ") " _
		& "order by clinicaltrialname "
else
	sQuery = sQuery & "where c.clinicaltrialid = " & lStudyId
end if

rsResult.open sQuery,Connect

'*************************
' For Each Study
'*************************
do until rsResult.eof 

WriteGroupHeader "Study", rsResult("clinicaltrialname") 

'*************************************************
' Get changed (site) data for server planned sdv's
'*************************************************
' general select
'sQuSelect = "Select mimessage.*, localidentifier1, visitname, crftitle, dataitemname, responsevalue "
sQuSelect = "Select trialsubject.trialsite, localidentifier1, trialsubject.personid, visitorder, visitcyclenumber,crfpageorder, crfpagecyclenumber,dataitemname, mimessage.*, visitname, crftitle, responsevalue "
sQuSelect = sQuSelect & "from mimessage, trialsubject, studyvisit,crfpage,dataitemresponsehistory,dataitem "
' general where
sQuWhere = "where DataItemResponseHistory.clinicaltrialid = trialsubject.clinicaltrialid "
sQuWhere = sQuWhere & "and DataItemResponseHistory.trialsite = trialsubject.trialsite "
sQuWhere = sQuWhere & "and DataItemResponseHistory.personid = trialsubject.personid "
sQuWhere = sQuWhere & "and DataItemResponseHistory.clinicaltrialid = crfpage.clinicaltrialid "
sQuWhere = sQuWhere & "and DataItemResponseHistory.crfpageid = crfpage.crfpageid "
sQuWhere = sQuWhere & "and DataItemResponseHistory.clinicaltrialid = studyvisit.clinicaltrialid "
sQuWhere = sQuWhere & "and DataItemResponseHistory.visitid = studyvisit.visitid "
sQuWhere = sQuWhere & "and DataItemResponseHistory.clinicaltrialid = dataitem.clinicaltrialid "
sQuWhere = sQuWhere &  "and DataItemResponseHistory.dataitemid = dataitem.dataitemid "
sQuWhere = sQuWhere & "and trialsubject.clinicaltrialid =  (select clinicaltrialid from clinicaltrial where clinicaltrialname = mimessagetrialname) "
sQuWhere = sQuWhere & "and trialsubject.trialsite = mimessage.mimessagesite "
sQuWhere = sQuWhere & "and trialsubject.personid = mimessage.mimessagepersonid "
sQuWhere = sQuWhere & "and mimessage.mimessagetype = 3 " ' sdv 
sQuWhere = sQuWhere & "and mimessage.mimessagesource = 0 " ' server
sQuWhere = sQuWhere & "and mimessage.mimessagehistory = 0 " ' current
sQuWhere = sQuWhere & "and mimessage.mimessagestatus = 2 " ' done
if sSite <> "" then
	sQuWhere = sQuWhere & "and mimessage.mimessagesite = '" & sSite & "' "
end if
if lSubjectId <> 0 then
	sQuWhere = sQuWhere & "and mimessage.mimessagepersonid = " & lSubjectId & " "
end if
' general order by
'sQuOrder = "order by trialsubject.trialsite,trialsubject.personid,localidentifier1,visitorder, "  
'sQuOrder = sQuOrder & "visitname,crfpageorder,crftitle,dataitemname "
sQuOrder = ""
' visit clause
sQuVisit = "and DataItemResponseHistory.visitid = mimessage.mimessagevisitid " & _
	"and DataItemResponseHistory.VisitCycleNumber = MIMessage.MIMessageVisitCycle "
' eform clause
sQuEform = "and DataItemResponseHistory.crfpagetaskid = mimessage.mimessagecrfpagetaskid "
' response clause
sQuResponse = "and DataItemResponseHistory.responsetaskid = mimessage.mimessageresponsetaskid "
' trialname
sQuTrialname = "and mimessagetrialname = '" & rsResult("clinicaltrialname") & "' "
if sDatabaseType = 1 then
	' sql server
	sQuNullFunc = "ISNULL"
else
	' oracle 
	sQuNullFunc = "NVL"
end if

'' put together large union query
'' response SDV data
'sQuery = sQuSelect & sQuWhere & sQuVisit & sQuEform & sQuResponse & sQuTrialname
'sQuery = sQuery & "and mimessage.mimessagescope = 4 " ' question
'sQuery = sQuery & "and mimessage.mimessageresponsetimestamp < DataItemResponseHistory.responsetimestamp " ' not same response
'' eform SDV data
'sQuery = sQuery & " UNION ALL "
'sQuery = sQuery & sQuSelect & sQuWhere & sQuVisit & sQuEform & sQuTrialname
'sQuery = sQuery & "and mimessage.mimessagescope = 3 " ' eform
'sQuery = sQuery & "and ( (mimessage.mimessagecreated < DataItemResponseHistory.responsetimestamp) " ' not same response
'sQuery = sQuery & "or (mimessage.mimessagecreated < " & sQuNullFunc & "(DataItemResponseHistory.importtimestamp,32874)) ) " ' sdv before import change
'' visit SDV data
'sQuery = sQuery & " UNION ALL "
'sQuery = sQuery & sQuSelect & sQuWhere & sQuVisit & sQuTrialname
'sQuery = sQuery & "and mimessage.mimessagescope = 2 " ' visit
'sQuery = sQuery & "and ( (mimessage.mimessagecreated < DataItemResponseHistory.responsetimestamp) " ' not same response
'sQuery = sQuery & "or (mimessage.mimessagecreated < " & sQuNullFunc & "(DataItemResponseHistory.importtimestamp,32874)) ) " ' sdv before import change
'' subject SDV data
'sQuery = sQuery & " UNION ALL "
'sQuery = sQuery & sQuSelect & sQuWhere & sQuTrialname
'sQuery = sQuery & "and mimessage.mimessagescope = 1 " ' subject
'sQuery = sQuery & "and ( (mimessage.mimessagecreated < DataItemResponseHistory.responsetimestamp) " ' not same response
'sQuery = sQuery & "or (mimessage.mimessagecreated < " & sQuNullFunc & "(DataItemResponseHistory.importtimestamp,32874)) ) " ' sdv before import change

'TA 2/9/04: no longer check on responsetimestamp so server entered data not included
' put together large union query
' response SDV data
sQuery = sQuSelect & sQuWhere & sQuVisit & sQuEform & sQuResponse & sQuTrialname
sQuery = sQuery & "and mimessage.mimessagescope = 4 " ' question
' sdv before import change
'convert null to 0 so that responses entered on server won't be included
sQuery = sQuery & "AND mimessage.mimessagecreated < " & sQuNullFunc & "(DataItemResponseHistory.importtimestamp,0)"
' eform SDV data
sQuery = sQuery & " UNION ALL "
sQuery = sQuery & sQuSelect & sQuWhere & sQuVisit & sQuEform & sQuTrialname
sQuery = sQuery & "and mimessage.mimessagescope = 3 " ' eform
' sdv before import change
'convert null to 0 so that responses entered on server won't be included
sQuery = sQuery & "AND mimessage.mimessagecreated < " & sQuNullFunc & "(DataItemResponseHistory.importtimestamp,0)"
' visit SDV data
sQuery = sQuery & " UNION ALL "
sQuery = sQuery & sQuSelect & sQuWhere & sQuVisit & sQuTrialname
sQuery = sQuery & "and mimessage.mimessagescope = 2 " ' visit
' sdv before import change
'convert null to 0 so that responses entered on server won't be included
sQuery = sQuery & "AND mimessage.mimessagecreated < " & sQuNullFunc & "(DataItemResponseHistory.importtimestamp,0)"
' subject SDV data
sQuery = sQuery & " UNION ALL "
sQuery = sQuery & sQuSelect & sQuWhere & sQuTrialname
sQuery = sQuery & "and mimessage.mimessagescope = 1 " ' subject
' sdv before import change
'convert null to 0 so that responses entered on server won't be included
sQuery = sQuery & "AND mimessage.mimessagecreated < " & sQuNullFunc & "(DataItemResponseHistory.importtimestamp,0)"

'response.write sQuery
'response.Flush()

rsResult1.open sQuery,Connect
	if rsResult1.eof then
		WritePara "No Done SDVs with changed data."
	else
		WriteTableStart
		WriteTableRowStart
		WriteHeaderCell "Study"
		WriteHeaderCell "Site"
		WriteHeaderCell "Subject"
		WriteHeaderCell "Visit"
		WriteHeaderCell "eForm"
		WriteHeaderCell "Question"
		WriteHeaderCell "Response"
		WriteHeaderCell "SDV Message"
		WriteHeaderCell "SDV Scope"
		WriteTableRowEnd

		do until rsResult1.eof 
'			if lSubjectId > 0 and lSubjectId = clng(rsResult1("mimessagepersonid")) then
			if lSubjectId = 0 or lSubjectId = clng(rsResult1("mimessagepersonid")) then
				WriteTableRowStart
				WriteCell rsResult1("mimessagetrialname")
				WriteCell rsResult1("mimessagesite")
				WriteCell fIdOrLabel(rsResult1("mimessagepersonid"), rsResult1("localidentifier1"))
				WriteCell rsResult1("visitname")
				WriteCell rsResult1("crftitle")
				WriteCell rsResult1("dataitemname")
				WriteCell rsResult1("responsevalue")
				WriteCell rsResult1("mimessagetext")
				' display scope text
				select case CInt(rsResult1("mimessagescope"))
					case 0:	WriteCell "Study"
					case 1: WriteCell "Subject"
					case 2: WriteCell "Visit"
					case 3: WriteCell "EForm"
					case 4: WriteCell "Question"
				end select		
				WriteTableRowEnd
			end if
			rsResult1.movenext
		loop

		WriteTableEnd
	end if
	rsResult1.Close
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