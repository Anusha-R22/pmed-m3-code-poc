<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'********************************************************************************************************
' Written By:	AN
'
' Revisions
'	RS	12 June 2003	Added missing cint() to comparisons
'	DPH 9 Jan 2004		handle null RRServerType
'********************************************************************************************************


'*************************
' Header block
'*************************

sReportTitle = "Study details"

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


sQuery = "Select clinicaltrialname, clinicaltrialdescription,phaseid,trialtypeid,studydefinition.*,statusname, keywords, expectedrecruitment "
sQuery = sQuery & "from clinicaltrial,trialstatus, studydefinition  "
sQuery = sQuery & " where clinicaltrial.statusid = trialstatus.statusid "
sQuery = sQuery & "   and clinicaltrial.clinicaltrialid = studydefinition.clinicaltrialid "
sQuery = sQuery & "   and clinicaltrial.clinicaltrialid = " & nClinicalTrialId

rsResult.open sQuery,Connect

if cint(rsResult("phaseid")) > 0 then
	 sQuery = "Select phasename "
	 sQuery = sQuery & "from trialphase "
	 sQuery = sQuery & " where phaseid = " & rsResult("phaseid")
	 rsResult1.open sQuery,Connect
	 sPhaseName = rsResult1("phasename")
	 rsResult1.close
else 
		 sPhaseName = ""
end if

if cint(rsResult("trialtypeid")) > 0 then
	 sQuery = "Select trialtypename "
	 sQuery = sQuery & "from trialtype "
	 sQuery = sQuery & " where trialtypeid = " & rsResult("trialtypeid")
	 rsResult1.open sQuery,Connect
	 sTrialTypeName = rsResult1("trialtypename")
	 rsResult1.close
else 
		 sTrialTypeName = ""
end if

WriteGroupHeader "Study", rsResult("clinicaltrialname")

WriteTableStart
WriteTableRowStart
WriteCell "Description:"
WriteCell rsResult("clinicaltrialdescription")
WriteTableRowEnd
WriteTableRowStart
WriteCell "Keywords:"
WriteCell rsResult("keywords")
WriteTableRowEnd
WriteTableRowStart
WriteCell "Phase:"
WriteCell sPhaseName
WriteTableRowEnd
WriteTableRowStart
WriteCell "Status:"
WriteCell rsResult("statusname")
WriteTableRowEnd
WriteTableRowStart
WriteCell "Type:"
WriteCell strialtypename
WriteTableRowEnd
WriteTableRowStart
WriteCell "Expected recruitment:"
WriteCell rsResult("expectedrecruitment")
WriteTableRowEnd
WriteTableRowStart
WriteCell "Last updated:"
WriteCell fConvertDate( rsResult("studydefinitiontimestamp") )
WriteTableRowEnd
WriteTableRowStart
WriteCell "Updated by:"
WriteCell rsResult("username")
WriteTableRowEnd
WriteTableRowStart
WriteCell "Subject label:"
WriteCell rsResult("trialsubjectlabel")
WriteTableRowEnd
if rsResult("DOBExpr") > "" then
	 WriteTableRowStart
	 WriteCell "Date of birth expression:"
	 WriteCell rsResult("DOBExpr") 
	 WriteTableRowEnd
end if
if rsResult("GenderExpr") > "" then
	 WriteTableRowStart
	 WriteCell "Gender expression:"
	 WriteCell rsResult("GenderExpr") 
	 WriteTableRowEnd
end if
if cint(rsResult("localtrialsubjectlabel")) > 0 then
	 WriteTableRowStart
	 WriteCell ""
	 WriteCell fLocal (rsResult("localtrialsubjectlabel") )
	 WriteTableRowEnd
end if
WriteTableRowStart
WriteCell "Standard date format:"
WriteCell rsResult("standarddateformat")
WriteTableRowEnd
WriteTableRowStart
WriteCell "Standard time format:"
WriteCell rsResult("standardtimeformat")
WriteTableRowEnd
if rsResult("CTCSchemeCode") > "" then
	 WriteTableRowStart
	 WriteCell "CTC Scheme:"
	 WriteCell rsResult("CTCSchemeCode")
	 WriteTableRowEnd
end if
WriteTableRowStart
WriteCell "Questions:"
if cint(rsResult("SingleUseDataItems")) = 0 then
	 WriteCell "Questions may be used on multiple eForms"
else
	 WriteCell "Questions can only be used on one eForm"
end if
WriteTableRowEnd
if rsResult("AREZZOMEMORY") > "" then
	 WriteTableRowStart
	 WriteCell "AREZZO memory settings:"
	 WriteCell rsResult("AREZZOMEMORY")
	 WriteTableRowEnd
end if
WriteTableRowStart
WriteCell "Registration server:"
nRRServerType = rsResult("RRServerType")
if isnull(nRRServerType) then
	nRRServerType = 0 'none
end if
WriteCell fRegistrationServer( nRRServerType )
WriteTableRowEnd
if cint(nRRServerType) = 3 then
	 WriteTableRowStart
	 WriteCell "HTTP address:"
	 WriteCell  rsResult("RRHTTPAddress") 
	 WriteTableRowEnd
	 WriteTableRowStart
	 WriteCell "User name:"
	 WriteCell rsResult("RRUserName") 
	 WriteTableRowEnd
	 WriteTableRowStart
	 WriteCell "Password:"
	 WriteCell "*********"
	 WriteTableRowEnd
	 WriteTableRowStart
	 WriteCell "Proxy server:"
	 WriteCell rsResult("RRProxyServer") 
	 WriteTableRowEnd
end if
WriteTableEnd

rsResult.Close

WriteTableStart
WriteTableRowStart
WriteHeaderCell "Object"
WriteHeaderCell "Number of items"
WriteTableRowEnd

sQuery =  "select count(visitid) as NumberOfItems from studyvisit "
sQuery = sQuery & "  where clinicaltrialid = " & request.querystring("clinicaltrialid")

rsResult.open sQuery,Connect

WriteTableRowStart
WriteCell "Visits"
WriteCell rsResult("NumberOfItems")
WriteTableRowEnd
rsResult.Close

sQuery =  "select count(crfpageid) as NumberOfItems from crfpage "
sQuery = sQuery & "  where clinicaltrialid = " & request.querystring("clinicaltrialid")

rsResult.open sQuery,Connect

WriteTableRowStart
WriteCell "Unique eCRF pages"
WriteCell rsResult("NumberOfItems")
WriteTableRowEnd
rsResult.Close

sQuery =  "select count(crfpageid) as NumberOfItems from studyvisitcrfpage "
sQuery = sQuery & "  where clinicaltrialid = " & request.querystring("clinicaltrialid")

rsResult.open sQuery,Connect

WriteTableRowStart
WriteCell "Total eCRF pages"
WriteCell rsResult("NumberOfItems")
WriteTableRowEnd
rsResult.Close

sQuery =  "select count(dataitemid) as NumberOfItems from dataitem "
sQuery = sQuery & "  where clinicaltrialid = " & request.querystring("clinicaltrialid")

rsResult.open sQuery,Connect

WriteTableRowStart
WriteCell "Questions"
WriteCell rsResult("NumberOfItems")
WriteTableRowEnd
rsResult.Close


sQuery =  "select count(dataitemid) as NumberOfItems from studyvisitcrfpage,crfelement "
sQuery = sQuery & "  where studyvisitcrfpage.clinicaltrialid = " & request.querystring("clinicaltrialid")
sQuery = sQuery & "    and studyvisitcrfpage.clinicaltrialid = crfelement.clinicaltrialid "
sQuery = sQuery & "    and studyvisitcrfpage.crfpageid = crfelement.crfpageid "
sQuery = sQuery & "    and crfelement.dataitemid > 0 "

rsResult.open sQuery,Connect


WriteTableRowStart
WriteCell "Estimated total questions per subject"
WriteCell rsResult("NumberOfItems")
WriteTableRowEnd
rsResult.Close

sQuery =  "select count(validationid) as NumberOfItems from dataitemvalidation "
sQuery = sQuery & "  where clinicaltrialid = " & request.querystring("clinicaltrialid")

rsResult.open sQuery,Connect

WriteTableRowStart
WriteCell "Validation checks"
WriteCell rsResult("NumberOfItems")
WriteTableRowEnd
rsResult.Close

WriteTableEnd



'*************************
' Subject numbering
'*************************

sQuery = "Select subjectnumbering.*, visitname, crftitle "
sQuery = sQuery & "from subjectnumbering,studyvisit,crfpage "
sQuery = sQuery & " where subjectnumbering.clinicaltrialid = " & nClinicalTrialId
sQuery = sQuery & "   and subjectnumbering.clinicaltrialid = studyvisit.clinicaltrialid "
sQuery = sQuery & "   and subjectnumbering.triggervisitid = studyvisit.visitid "
sQuery = sQuery & "   and subjectnumbering.clinicaltrialid = crfpage.clinicaltrialid "
sQuery = sQuery & "   and subjectnumbering.triggerformid = crfpage.crfpageid "
sQuery = sQuery & "   and subjectnumbering.useregistration > 0 "

rsResult.open sQuery,Connect

if rsResult.eof then
	 ' Write nothing
else
		WriteTableStart
	 WriteTableRowStart
	 WriteCell "Registration occurs after visit:"
	 WriteCell rsResult("visitname") 
	 WriteTableRowEnd
	 WriteTableRowStart
	 WriteCell "Registration occurs after eForm:"
	 WriteCell rsResult("crftitle") 
	 WriteTableRowEnd
	 WriteTableRowStart
	 WriteCell "Start number:"
	 WriteCell rsResult("startnumber") 
	 WriteTableRowEnd
	 WriteTableRowStart
	 WriteCell "Maximum width:"
	 if clng(rsResult("numberwidth")) > 0 then
	 	 WriteCell rsResult("numberwidth")
	 else
		 WriteCell "Unlimited"
	 end if 
	 WriteTableRowEnd
	 WriteTableRowStart
	 WriteCell "Subject identifier prefix:"
	 WriteCell rsResult("prefix") 
	 WriteTableRowEnd
	 WriteTableRowStart
	 WriteCell "Restart numbering after change in prefix:"
	 if cint(rsResult("useprefix")) = 0 then
	 	 WriteCell "No"
	 else
	 	 WriteCell "Yes"
	 end if 
	 WriteTableRowEnd
	 WriteTableRowStart
	 WriteCell "Subject identifier suffix:"
	 WriteCell rsResult("suffix") 
	 WriteTableRowEnd
	 WriteTableRowStart
	 WriteCell "Restart numbering after change in suffix:"
	 if cint(rsResult("usesuffix")) = 0 then
	 	 WriteCell "No"
	 else
	 	 WriteCell "Yes"
	 end if 
	 WriteTableRowEnd
	 WriteTableEnd

end if

rsResult.Close

'*************************
' Registrationchecks
'*************************

sQuery = "Select * "
sQuery = sQuery & "from eligibility "
sQuery = sQuery & " where clinicaltrialid = " & nClinicalTrialId
sQuery = sQuery & " order by eligibilitycode "

rsResult.open sQuery,Connect

if rsResult.eof then
	 ' Write nothing
else
		WriteTableStart
		WriteTableRowStart
		WriteHeaderCell "Registration code"
		WriteHeaderCell "Registration conditions"
		WriteTableRowEnd
		do until rsResult.eof 
			 WriteTableRowStart
			 WriteCell rsResult("eligibilitycode")
			 WriteCell rsResult("condition")
			 WriteTableRowEnd
			 rsResult.movenext
		loop
		WriteTableEnd
end if

rsResult.Close

'*************************
' Uniqueness checks
'*************************

sQuery = "Select * "
sQuery = sQuery & "from uniquenesscheck "
sQuery = sQuery & " where clinicaltrialid = " & nClinicalTrialId
sQuery = sQuery & " order by checkcode "

rsResult.open sQuery,Connect

if rsResult.eof then
	 ' Write nothing
else
		WriteTableStart
		WriteTableRowStart
		WriteHeaderCell "Uniqueness code"
		WriteHeaderCell "Uniqueness expressions"
		WriteTableRowEnd
		do until rsResult.eof 
			 WriteTableRowStart
			 WriteCell rsResult("checkcode")
			 WriteCell rsResult("expression")
			 WriteTableRowEnd
			 rsResult.movenext
		loop
		WriteTableEnd
end if

rsResult.Close



'*************************
' Reasons for change
'*************************

sQuery = "Select * "
sQuery = sQuery & "from reasonforchange "
sQuery = sQuery & " where clinicaltrialid = " & nClinicalTrialId
sQuery = sQuery & "   and reasontype = 0"
sQuery = sQuery & " order by reasonforchange "

rsResult.open sQuery,Connect

if rsResult.eof then
	 WritePara "<b>No reasons for change have been set up for this study.</b>"
else
		WriteTableStart
		WriteTableRowStart
		WriteHeaderCell "Reasons for change"
		WriteTableRowEnd
		do until rsResult.eof 
			 WriteTableRowStart
			 WriteCell rsResult("reasonforchange")
			 WriteTableRowEnd
			 rsResult.movenext
		loop
		WriteTableEnd
end if

rsResult.Close

'*************************
' Reasons for overrule
'*************************

sQuery = "Select * "
sQuery = sQuery & "from reasonforchange "
sQuery = sQuery & " where clinicaltrialid = " & nClinicalTrialId
sQuery = sQuery & "   and reasontype = 1"
sQuery = sQuery & " order by reasonforchange "

rsResult.open sQuery,Connect

if rsResult.eof then
	 WritePara "<b>No reasons for overrule have been set up for this study.</b>"
else
		WriteTableStart
		WriteTableRowStart
		WriteHeaderCell "Reasons for overrule"
		WriteTableRowEnd
		do until rsResult.eof 
			 WriteTableRowStart
			 WriteCell rsResult("reasonforchange")
			 WriteTableRowEnd
			 rsResult.movenext
		loop
		WriteTableEnd
end if

rsResult.Close

'*************************
' Study documents
'*************************

sQuery = "Select * "
sQuery = sQuery & "from studydocument "
sQuery = sQuery & " where clinicaltrialid = " & nClinicalTrialId

rsResult.open sQuery,Connect

if rsResult.eof then
	 WritePara "<b>There are no reference documents attached to this study.</b>"
else
		WriteTableStart
		WriteTableRowStart
		WriteHeaderCell "Study documents"
		WriteTableRowEnd
		do until rsResult.eof 
			 WriteTableRowStart
			 WriteCell rsResult("DocumentPath")
			 WriteTableRowEnd
			 rsResult.movenext
		loop
		WriteTableEnd
end if

rsResult.Close





'*************************
' Study status history
'*************************

sQuery = "Select trialstatuschangeid,username,statuschangedtimestamp,statuschangedtimestamp_TZ, statusname "
sQuery = sQuery & "from trialstatushistory,trialstatus "
sQuery = sQuery & " where clinicaltrialid = " & nClinicalTrialId
sQuery = sQuery & "   and trialstatushistory.statusid = trialstatus.statusid "
sQuery = sQuery & " order by trialstatuschangeid "

rsResult.open sQuery,Connect

if rsResult.eof then
	 WritePara "<b>No status changes have been made to this study.</b>"
else
		WriteTableStart
		WriteTableRowStart
		WriteHeaderCell "Status"
		WriteHeaderCell "Date / time"
		WriteHeaderCell "User"
		WriteTableRowEnd
		do until rsResult.eof 
			 WriteTableRowStart
			 WriteCell rsResult("statusname")
			 WriteCell fConvertDate(rsResult("statuschangedtimestamp"))
			 WriteCell rsResult("username")
			 WriteTableRowEnd
			 rsResult.movenext
		loop
		WriteTableEnd
end if

rsResult.Close

'*************************
' Study versions
'*************************

sQuery = "Select * "
sQuery = sQuery & "from studyversion "
sQuery = sQuery & " where clinicaltrialid = " & nClinicalTrialId

rsResult.open sQuery,Connect

if rsResult.eof then
	 WritePara "<b>No versions have been distributed.</b>"
else
		WriteTableStart
		WriteTableRowStart
		WriteHeaderCell "Version"
		WriteHeaderCell "Date / time"
		WriteHeaderCell "Description"
		WriteTableRowEnd
		do until rsResult.eof 
			 WriteTableRowStart
			 WriteCell rsResult("StudyVersion")
			 WriteCell fConvertDate(rsResult("VersionTimestamp"))
			 WriteCell rsResult("VersionDescription")
			 WriteTableRowEnd
			 rsResult.movenext
		loop
		WriteTableEnd
end if

rsResult.Close


set rsResult = Nothing
set rsResult1 = Nothing

'*************************
' Footer block
'*************************

%>
<!--#include file="report_footer.asp" -->
<!--#include file="report_close.asp" -->