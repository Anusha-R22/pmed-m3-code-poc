<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Data views"

sIncludeVML = 0 'Don't include VML styles

%>
<!--#include file="report_initialise.asp" -->
<!--#include file="report_open_macro_database.asp" -->
<!--#include file="report_functions.asp" -->
<!--#include file="report_header.asp" -->
<%

'*************************
' Content block
'*************************

nClinicalTrialId = request.querystring("clinicaltrialid")

Set rsResult = CreateObject("ADODB.Recordset")
Set rsResult1 = CreateObject("ADODB.Recordset")

Dim sTables

sTables = split(request.querystring("tablename"), ",")

sQuery = "Select "
if request.querystring("NumRows") > "" then
	 sQuery = sQuery & " top " & request.querystring("NumRows")
end if
	 sQuery = sQuery & " localidentifier1 "
	 
for n = lbound(sTables) to ubound(sTables)	 
	 	 sQuery = sQuery & ",t" & n & ".* "
next		 
	 sQuery = sQuery &  "	  from trialsubject " 
for n = lbound(sTables) to ubound(sTables)
	 sQuery = sQuery & ", " & sTables(n) & " t" & n 
next	 
sQuery = sQuery & " where "
for n = lbound(sTables) to ubound(sTables)
if n > lbound(sTables) then
sQuery = sQuery & " and "
end if
sQuery = sQuery & " trialsubject.clinicaltrialid = t" & n & ".clinicaltrialid "
sQuery = sQuery & " and  trialsubject.trialsite = t" & n & ".site "
sQuery = sQuery & " and  trialsubject.personid = t" & n & ".personid "
next	 

if request.querystring("SelectClause") > "" then
	 sQuery = sQuery & " and " & request.querystring("SelectClause")
end if
' dph 12/02/2004 - study/site permissions
sQuery = sQuery & " and " & replace(replace(sStudySiteSQL, "clinicaltrial.", "trialsubject."), "trialsite.", "trialsubject." )
sQuery = sQuery & " order by t0.site,localidentifier1, t0.visitid, t0.visitcyclenumber,t0.crfpageid,t0.crfpagecyclenumber "

on error resume next
rsResult1.open sQuery,Connect
if connect.errors.count > 0 then
response.write "SQL query is not valid"
else


WriteTableStart
WriteTableRowStart

for each oField in rsResult1.Fields

if oField.Name = "localidentifier1" then
			 WriteHeaderCell "Subject"
elseif oField.Name <> "ClinicalTrialId" and oField.Name <> "Site" and oField.Name <> "PersonId" and oField.Name <> "VisitId" and oField.Name <> "VisitCycleNumber" and oField.Name <> "CRFPageId" and oField.Name <> "CRFPageCycleNumber" then
			 WriteHeaderCell oField.Name 
end if

next
WriteTableRowEnd


do until rsResult1.eof 
WriteTableRowStart

for each oField in rsResult1.Fields
if oField.Name <> "ClinicalTrialId" and oField.Name <> "Site" and oField.Name <> "PersonId" and oField.Name <> "VisitId" and oField.Name <> "VisitCycleNumber" and oField.Name <> "CRFPageId" and oField.Name <> "CRFPageCycleNumber" then
	 WriteCell oField.Value 
end if

next
WriteTableRowEnd
rsResult1.movenext
loop

WriteTableEnd

end if

rsResult1.Close
set RsResult1 = Nothing

'*************************
' Footer block
'*************************

%>
<!--#include file="report_footer.asp" -->
<!--#include file="report_close.asp" -->