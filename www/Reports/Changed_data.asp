<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "Data values changed > 2 times"

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

Set rsResult = CreateObject("ADODB.Recordset")

' RS 07JUN2005: Optimized SQL by using a subquery instead of many linked tables
sQuery = "select dir.clinicaltrialid,clinicaltrial.clinicaltrialname,dir.trialsite,trialsubject.localidentifier1,studyvisit.visitcode," & vbNewLine & _
"       crfpage.crfpagecode, dataitem.dataitemcode,  dir.changecount" & vbNewLine & _
"from" & vbNewLine & _
"(" & vbNewLine & _
"Select clinicaltrialid," & vbNewLine & _
"       TrialSite," & vbNewLine & _
"       PersonID," & vbNewLine & _
"       VisitID," & vbNewLine & _
"       CRFPageID," & vbNewLine & _
"       DataItemID," & vbNewLine & _
"       ChangeCount" & vbNewLine & _
"  from dataitemresponse" & vbNewLine & _
" where ChangeCount > 2 and" & vbNewLine & replace(replace(sStudySiteSQL,"clinicaltrial.",""),"trialsite.","") & _
" ) dir, trialsubject, studyvisit, crfpage, dataitem, clinicaltrial" & vbNewLine & _
" where trialsubject.clinicaltrialid = dir.clinicaltrialid and trialsubject.trialsite = dir.trialsite and trialsubject.personid = dir.personid" & vbNewLine & _
" and clinicaltrial.clinicaltrialid = dir.clinicaltrialid" & vbNewLine & _
" and studyvisit.clinicaltrialid = dir.clinicaltrialid and studyvisit.visitid = dir.visitid" & vbNewLine & _
" and crfpage.clinicaltrialid = dir.clinicaltrialid and crfpage.crfpageid = dir.crfpageid" & vbNewLine & _
" and dataitem.clinicaltrialid = dir.clinicaltrialid and dataitem.dataitemid = dir.dataitemid" & vbNewLine & _
"order by dir.Clinicaltrialid,dir.ChangeCount desc"

'response.write sQuery
response.flush()

rsResult.open sQuery,Connect

curStudy = ""
do until rsResult.eof 
	if rsResult("clinicaltrialname") <> curStudy then
		if curStudy<>"" then WriteTableEnd

		WriteGroupHeader "Study", rsResult("clinicaltrialname") 

		WriteTableStart
		WriteTableRowStart
		WriteHeaderCell "Site"
		WriteHeaderCell "Subject"
		WriteHeaderCell "Visit"
		WriteHeaderCell "eForm"
		WriteHeaderCell "Question"
		WriteHeaderCell "Number of changes"
		WriteTableRowEnd
		
		curStudy = rsResult("clinicaltrialname") 
	end if
	WriteTableRowStart
	WriteCell rsResult("TrialSite") 
	WriteCell rsResult("LocalIdentifier1") 
	WriteCell rsResult("VisitCode") 
	WriteCell rsResult("CRFPageCode")
	WriteCell rsResult("DataItemCode" ) 
	WriteCentredCell rsResult("ChangeCount") 
	WriteTableRowEnd
	rsResult.movenext
loop

WriteTableEnd

rsResult.close
set RsResult = Nothing



'*************************
' Footer block
'*************************

%>
  <!--#include file="report_footer.asp" -->
<!--#include file="report_close.asp" -->