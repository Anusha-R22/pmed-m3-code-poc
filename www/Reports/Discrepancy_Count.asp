<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Discrepancy Count"

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

sQuery = "Select distinct c.clinicaltrialid,clinicaltrialname "
sQuery = sQuery & "from clinicaltrial c  "
sQuery = sQuery & " where c.clinicaltrialid in (" & sPermittedStudyList & ") "
sQuery = sQuery & "order by clinicaltrialname "

rsResult.open sQuery,Connect

'*************************
' For Each Study
'*************************
do until rsResult.eof 

WriteGroupHeader "Study", rsResult("clinicaltrialname") 

'*************************
' Get Discrepancy Count from MACRO Object
'*************************
sQuery = "select MIMessageTrialname, MIMessageSite, MIMessageType, MIMessageStatus, count(MIMessageID) DiscrepancyCount"
sQuery = sQuery & " from mimessage"
sQuery = sQuery & " where mimessageid in "
sQuery = sQuery & " ( select max(mimessageid) as maxid from mimessage group by mimessageobjectid"
sQuery = sQuery & " )"
sQuery = sQuery & " and mimessagetype = 0"
sQuery = sQuery & " and MIMessageTrialName = '" & rsResult("clinicaltrialname") & "'"
sQuery = sQuery & " group by MIMessageTrialname, MIMessageSite, MIMessageType, MIMessageStatus"
sQuery = sQuery & " order by MIMessageTrialname, MIMessageSite, MIMessageType, MIMessageStatus"


'*************************
' RS 06JUN2005
' The original query did not retrieve correct results in some cases. This new query uses the MIMessageHistory column
' to retrieve only current values (instead of subquery)
'*************************
sQuery = "select MIMessageTrialname," & vbNewLine & _
"       MIMessageSite," & vbNewLine & _
"       MIMessageType," & vbNewLine & _
"       MIMessageStatus," & vbNewLine & _
"       count(MIMessageID) DiscrepancyCount" & vbNewLine & _
"  from mimessage t" & vbNewLine & _
" where mimessagehistory=0" & vbNewLine & _
"       and MIMessageTrialName = '" & rsResult("clinicaltrialname") & "'" & vbNewLine & _
"       and MIMessagetype = 0" & vbNewLine & _
" group by MIMessageTrialname, MIMessageSite, MIMessageType, MIMessageStatus" & vbNewLine & _
" order by MIMessageTrialname, MIMessageSite, MIMessageType, MIMessageStatus"



'sQuery = "select mimessagesite,mimessagestatus,count(*) as TotMsg from mimessage where MIMessageType = 0"
'sQuery = sQuery & " and MIMessageTrialName = '" & rsResult("clinicaltrialname") & "'"
'sQuery = sQuery & " and MIMessageHistory=0 "
'sQuery = sQuery & " group by mimessagesite,mimessagestatus "

'response.write sQuery
'response.Flush()

rsResult1.open sQuery,Connect
	if rsResult1.eof then
		WritePara "No Discrepancies."
	else
		WriteTableStart
		WriteTableRowStart
		WriteHeaderCell "Site"
		WriteHeaderCell "Status"
		WriteHeaderCell "Discrepancies"
		WriteHeaderCell "Details"
		WriteTableRowEnd

		do while not rsResult1.eof 
			 WriteTableRowStart
			 WriteCell rsResult1("mimessagesite")
			 select case CInt(rsResult1("MIMessageStatus"))
			 	case 0:	WriteCell("Raised")
				case 1: WriteCell("Responded")
				case 2: WriteCell("Closed")
			 end select
			 WriteCell rsResult1("DiscrepancyCount")
			 ' DPH - 15/03/2004 - use WriteLink to avoid problms in CSV/Excel 
			 sQueryString = "clinicaltrialname=" & rsResult("clinicalTrialName")
			 sQueryString = sQueryString & "&trialsite=" & rsResult1("mimessagesite")
			 sQueryString = sQueryString & "&mimessagestatus=" & rsResult1("MIMessageStatus")
			 WriteLink "Details", "Discrepancy_Details.asp", sQueryString
			 'sLink = "<A HREF=""Discrepancy_Details.asp?clinicaltrialname=" & rsResult("clinicalTrialName")
			 'sLink = sLink & "&trialsite=" & rsResult1("mimessagesite")
			 'sLink = sLink & "&mimessagestatus=" & rsResult1("MIMessageStatus")
			 'sLink = sLink & """>Details</A>"
			 'WriteCell sLink
			 WriteTableRowEnd
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