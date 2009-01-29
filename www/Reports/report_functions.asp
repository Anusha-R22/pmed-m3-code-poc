<% 
'********************************************************************************************************
' Written By:	AN
'
' Revisions
'	RS	12 June 2003	Added missing cint() to comparisons/select case
'********************************************************************************************************
sub WriteTableStart
if sReportType = 0 then
response.write "<table cellspacing=""8px"">"
elseif sReportType = 1 then
response.write "<table>"
elseif sReportType = 2 then
end if 
end sub

sub WriteTableEnd
if sReportType = 0 then
response.write "</table>"
elseif sReportType = 1 then
response.write "</table>"
elseif sReportType = 2 then
end if 
end sub

sub WriteTableRowStart
if sReportType = 0 then
response.write "<tr valign=""top"">"
elseif sReportType = 1 then
response.write "<tr>"
elseif sReportType = 2 then
end if 
end sub

sub WriteTableRowEnd
if sReportType = 0 then
response.write "</tr>"
elseif sReportType = 1 then
response.write "</tr>"
elseif sReportType = 2 then
response.write chr(10) 
end if 
end sub

sub WriteHeaderCell (sContent)
if sReportType = 0 then
response.write "<th valign=""top"">" & sContent & "</th>"
elseif sReportType = 1 then
response.write "<th>" & sContent & "</th>"
elseif sReportType = 2 then
response.write "," & sContent
end if 
end sub

sub WriteHeaderLink (sContent)
if sReportType = 0 then
response.write "<th>" & sContent & "</th>"
elseif sReportType = 1 then
' RS 10/06/2003: Added (block was empty)
response.write "<th>" & sContent & "</th>"
elseif sReportType = 2 then
end if 
end sub

sub WriteCell (sContent)
if sReportType = 0 then
response.write "<td>" & sContent & "</td>"
elseif sReportType = 1 then
response.write "<td>" & sContent & "</td>"
elseif sReportType = 2 then
response.write "," & sContent
end if 
end sub

sub WriteTableCellStart (sProperties)
if sReportType = 0 then
response.write "<td " & sProperties & ">" 
elseif sReportType = 1 then
response.write "<td>" 
elseif sReportType = 2 then
end if 
end sub

sub WriteTableCellEnd ()
if sReportType = 0 then
response.write "</td>" 
elseif sReportType = 1 then
response.write "</td>" 
elseif sReportType = 2 then
end if 
end sub

sub WritePara (sContent)
if sReportType = 0 then
response.write "<p>" & sContent & "</p>"
elseif sReportType = 1 then
response.write "<p>" & sContent & "</p>"
elseif sReportType = 2 then
response.write sContent & chr(10)
end if 
end sub

sub WriteFixedWidthCell (sContent, nWidth)
if sReportType = 0 then
response.write "<td width=" & nWidth & ">" & sContent & "</td>"
elseif sReportType = 1 then
response.write "<td>" & sContent & "</td>"
elseif sReportType = 2 then
response.write "," & sContent
end if 
end sub

sub WriteCentredCell (sContent)
if sReportType = 0 then
response.write "<td align=""center"">" & sContent & "</td>"
elseif sReportType = 1 then
response.write "<td>" & sContent & "</td>"
elseif sReportType = 2 then
response.write "," & sContent
end if 
end sub

sub WriteLink (sDescription, sURL, sQueryString)
' dph 15/03/2004 - write link description in csv/excel
if sReportType = 0 then
response.write "<td><a href=""" & sURL & "?" & sQueryString & """>" & sDescription & "</a></td>"
elseif sReportType = 1 then
	WriteCell sDescription
elseif sReportType = 2 then
	WriteCell sDescription
end if 
end sub


sub WriteGroupHeader (sGroup,sContent)
if sReportType = 0 then
response.write "<p class=""GroupHeader1"">" & sGroup & ": " & sContent & "</p>"
elseif sReportType = 1 then
response.write "<p class=""GroupHeader1"">" & sGroup & ": " & sContent & "</p>"
elseif sReportType = 2 then
response.write sGroup & "," & sContent & chr(10)
end if 
end sub

function fEnabled (sStatus)
select case cint(sStatus) 
case 0  
	 fEnabled =  "Enabled"
case 1
	 fEnabled =  "Disabled"
case else 
	 fEnabled =  "Unknown value"
end select

end function


function fRoleEnabled (sStatus)
select case cint(sStatus) 
case 1  
	 fRoleEnabled =  "Enabled"
case 0
	 fRoleEnabled =  "Disabled"
case else 
	 fRoleEnabled =  "Unknown value"
end select

end function

function fSiteLocation (sLocation)
select case cint(sLocation) 
case 0
	 fSiteLocation =  "Server / Web"
case 1
	 fSiteLocation =  "Remote"
case else 
	 fSiteLocation =  "Unknown value"
end select
end function

function fMACROOnly (sValue)
select case cint(sValue) 
case 0
	 fMACROOnly =  ""
case 1
	 fMACROOnly =  "(Do not send to Oracle Clinical)"
case else 
	 fMACROOnly =  "Unknown value"
end select
end function

function fActive (sValue)
select case cint(sValue) 
case 0
	 fActive =  "(Inactive)"
case 1
	 fActive =  ""
case else 
	 fActive =  "Unknown value"
end select
end function

function fCase (sValue)
select case cint(sValue) 
case 0
	 fCase =  "As entered"
case 1
	 fCase =  "Upper case"
case 2
	 fCase =  "Lower case"
case else 
	 fCase =  "Unknown value"
end select
end function

function fVisitRepeats (nRepeats)
if isnull(nRepeats) then
	fVisitRepeats =  "None"
else
	if cint(nRepeats) >= -1 then
		select case cint(nRepeats) 
		case 0
	 		fVisitRepeats =  "None"
		case -1
			fVisitRepeats =  "Unlimited"
		case else 
			fVisitRepeats =  nRepeats
		end select
	else
		fVisitRepeats = "1"
	end if
end if
end function

function fFormRepeats (nRepeats)
if cint(nRepeats) >= 0 then
	 select case cint(nRepeats) 
	 case 0
	 	 fFormRepeats =  ""
	 case 1
		 fFormRepeats =  "(Repeating)"
	 case else 
		 fFormRepeats =  "Unknown value"
	 end select
else
	 fFormRepeats = ""
end if
end function

function fVisiteForm (nValue)
if cint(nValue) >= 0 then
	 select case cint(nValue) 
	 case 0
	 	 fVisiteForm =  ""
	 case 1
		 fVisiteForm =  "(Visit eForm)"
	 case else 
		 fVisiteForm =  "Unknown value"
	 end select
else
	 fVisiteForm = ""
end if
end function

function fPageSize (nValue)
if isnull(nValue) then
	 fPageSize = "Portrait"
else
if clng(nValue) >= 0 then
	 select case clng(nValue)
	 case 8515
	 	 fPageSize =  "Portrait"
	 case 14500
		 fPageSize =  "Landscape"
	 case else 
		 fPageSize =  nValue
	 end select
else
	 fPageSize = "Portrait"
end if
end if
end function

function fRegistrationServer (sServer)
if isnull(sServer) then
	 fRegistrationServer = "None"
else
	select case 	clng(sServer )
	case 0
		 fRegistrationServer =  "None"
	case 1
		 fRegistrationServer =  "Local"
	case 2
		 fRegistrationServer =  "Trial office"
	case 3
		 fRegistrationServer =  "Remote"
	case else 
		 fRegistrationServer =  "None"
	end select
end if
end function

function fIdOrLabel (sId, sLabel)
if sLabel > "" then
	 fIdOrLabel = sLabel
else
		if sReportType=1 then
			' RS 11/06/2003: Do not display brackets in EXCEL type report
			fIdOrLabel = sId
		else
			fIdOrLabel = "(" & sId & ")"
		end if
end if
end function

function fConvertDate (nDate)
		fConvertDate = cdate(nDate)
end function

function fLocal (nLocal)
select case nLocal 
case 0  
	 fLocal =  "Transferred to server from remote sites"
case 1
	 fLocal =  "Local value retained at remote sites and not transferred to server"
case else 
	 fLocal =  "Unknown value"
end select
end function

function fFormStatusImage (sStatus)

select case cint(sStatus) 
case -10
'	 fFormStatusImage =  "<img src=""" & server.mappath("../img/icof_new.gif) & """ />"
case 0
	 fFormStatusImage =  "<img src=""" & server.mappath("../img/icof_missing.gif") & """ />"
case 10
	 fFormStatusImage =  "<img src=""" & server.mappath("../img/icof_OK.gif") & """ />"
case 25
	 fFormStatusImage =  "<img src=""" & server.mappath("../img/icof_inform.gif") & """ />"
case 30
	 fFormStatusImage =  "<img src=""" & server.mappath("../img/icof_warn.gif") & """ />"
case else 
	 fFormStatusImage = sStatus
end select
end function

function fStatus (sStatus)
' DPH 17/03/2004 - Adding missing statuses
select case cint(sStatus) 
case -20
	 fStatus =  "Cancelled By User"
case -10
	 fStatus =  "New"
case -8
	 fStatus =  "Not Applicable"
case -5
	 fStatus =  "Unobtainable"
case 0
	 fStatus =  "OK"
case 10
	 fStatus =  "Missing"
case 20
	 fStatus =  "Inform"
case 25
	 fStatus =  "Ok Warning"
case 30
	 fStatus =  "Warning"
case else 
	 fStatus = "Unknown status"
end select
end function

' MLM 20/07/05: Added:
function HiddenFormElements(asVisibleElements)
	dim sForm
	dim asParameters
	dim sParameter
	dim lCount
	asParameters = Split(Request.QueryString, "&")
	for lCount = 0 to ubound(asParameters)
		sParameter = Split(asParameters(lCount), "=")(0)
		if not Contains(asVisibleElements, sParameter) then
			sForm = sForm & "<input type=hidden name='" & sParameter & "' value='" & Request.QueryString(sParameter) & "'>"
		end if
	next
	HiddenFormElements = sForm
end function

function Contains(asStrings, sString)
	for lCount = 0 to ubound(asStrings)
		if asStrings(lCount) = sString then
			Contains = true
			exit function
		end if
	next
	Contains = false
end function

' MLM 20/07/05: Added:
'	asStudies: An array of study names to be offered to the user, or an empty array for all the user's studies
'	bShowSiteList: Should the user be offered a choice of sites?
'	asSites: if(bShowSiteList): An array of site names to be offered to the user, or an empty array for all the user's sites
'	bShowSubjectList: if(bShowSiteList): Should the user be offered a choice of sites?
sub WriteSelectionHeader(asStudies, bShowSiteList, asSites, bShowSubjectList)
dim sSQL
dim lCount
dim rsSubjects
dim sFirstStudy
dim sFirstSite
dim sFirstSubject
dim sStudyHtml
dim sSiteHtml
dim sSubjectHtml
dim sScript
dim lRecordCount
dim bMultipleStudies
dim bMultipleSites
dim bMultipleSubjects
dim lStudyId
dim sSite
dim sSubject
dim sSubjectLabel
	if not bShowSiteList then bShowSubjectList = false
	'prepare query
	sSQL = "SELECT ClinicalTrialName, ClinicalTrial.ClinicalTrialId, TrialSite, LocalIdentifier1, PersonId" & _
		" FROM ClinicalTrial, TrialSubject WHERE ClinicalTrial.ClinicalTrialId = TrialSubject.ClinicalTrialId"
	if ubound(asStudies) > -1 then
		sSQL= sSQL & " AND ClinicalTrialName IN ("
		for lcount = 0 to ubound(asStudies)
			sSQL = sSQL & "'" & asStudies(lCount) & "',"
		next
		sSQL = Mid(sSQL, 1, len(sSQL) - 1) & ")"
	end if
	if ubound(asSites) > -1 then
		sSQL= sSQL & " AND TrialSite IN ("
		for lcount = 0 to ubound(asSites)
			sSQL = sSQL & "'" & asSites(lCount) & "',"
		next
		sSQL = Mid(sSQL, 1, len(sSQL) - 1) & ")"
	end if
	sSQL = sSQL & " AND " & replace(sStudySiteSQL, "trialsite.", "TrialSubject.")
	sSQL = sSQL & " ORDER BY  ClinicalTrialName, TrialSite, LocalIdentifier1"

	Set rsSubjects = CreateObject("ADODB.Recordset")
	rsSubjects.Open sSQL, Connect
	
	lRecordCount = 0
	if rsSubjects.EOF then
		'the report isn't going to be any use; give up
		Response.Write "This report contains no data."
		exit sub
	end if
	bMultipleStudies = false
	bMultipleSites = false
	bMultipleSubjects = false
	if Request.QueryString("study")= "" or not isnumeric(Request.QueryString("study")) then
		lStudyId = 0
	else
		lStudyId = clng(Request.QueryString("study"))
	end if

	do
		lRecordCount = lRecordCount + 1
		sScript = sScript & "data[" & lRecordCount & "]=Array('" & rsSubjects.Fields(0).Value & "', " & _
			rsSubjects.Fields(1).Value &  ", '" & _
			rsSubjects.Fields(2).Value & "', '"
		if isnull(rsSubjects.Fields(3).Value) then
			sScript = sScript &  rsSubjects.Fields(4).Value
		else
			sScript = sScript & replace(rsSubjects.Fields(3).Value, "'", "''")
		end if
		sScript = sScript & "', " & rsSubjects.Fields(4).Value & ");" & vbCrLf
		'study
		if instr(sStudyHtml, ">" & rsSubjects.Fields("ClinicalTrialName").Value & "<") = 0 then
			if sStudyHtml <> "" then
				bMultipleStudies = true
			end if
			sStudyHtml = sStudyHtml & "<option value='" & rsSubjects.Fields("ClinicalTrialId").Value & "'"
				if clng(rsSubjects.Fields("ClinicalTrialId").Value) = lStudyId then
					sStudyHtml = sStudyHtml & " selected"
				end if
			sStudyHtml = sStudyHtml & ">" & rsSubjects.Fields("ClinicalTrialName").Value & "</option>"
		end if
		'site
		if instr(sSiteHtml, ">" & rsSubjects.Fields("TrialSite").Value & "<") = 0 then
			if lStudyId = 0 or clng(rsSubjects.Fields("ClinicalTrialId").Value) = lStudyId then
				if sSiteHtml <> "" then
					bMultipleSites = true
				end if
				sSite = rsSubjects.Fields("TrialSite").Value
				sSiteHtml = sSiteHtml & "<option value=" & sSite
				if sSite = Request.QueryString("site") then
					sSiteHtml = sSiteHtml & " selected"
				end if
				sSiteHtml = sSiteHtml & ">" & sSite & "</option>"
			end if
		end if
		'subject
		if (lStudyId = 0 or clng(rsSubjects.Fields("ClinicalTrialId").Value) = lStudyId) _
			and (Request.QueryString("site") = "" or rsSubjects.Fields("TrialSite").Value = Request.QueryString("site")) then
			if sSubjectHtml <> "" then
				bMultipleSubjects = true
			end if
			sSubject = rsSubjects.Fields("ClinicalTrialId").Value & "`" & _
				rsSubjects.Fields("TrialSite").Value & "`" & _
				rsSubjects.Fields("PersonId").Value
			if isnull(rsSubjects.Fields("LocalIdentifier1").Value) then
				sSubjectLabel = rsSubjects.Fields("PersonId").Value
			else
				sSubjectLabel = rsSubjects.Fields("LocalIdentifier1").Value
			end if
			sSubjectHtml = sSubjectHtml & "<option value='" & sSubject & "'"
			if sSubject = Request.QueryString("subject") then
				sSubjectHtml = sSubjectHtml & " selected"
			end if
			sSubjectHtml = sSubjectHtml & ">" & sSubjectLabel & "</option>"
		end if

		rsSubjects.MoveNext
	loop until rsSubjects.EOF
	rsSubjects.MoveFirst
%>
	<script language=JavaScript>var data = new Array(<%=lRecordCount%>);
	<%=sScript%>
	function SetContents(oSelect, displayType, textCol, valueCol, lStudyId, sSite){
		var value = oSelect.options[oSelect.selectedIndex].value;
		oSelect.options.length = 0;
		oSelect.options[0] = new Option('All ' + displayType, '', false, false);
		for(var count=1;count<data.length;count++)
			if((lStudyId==''||data[count][1]==lStudyId) && (sSite==''||data[count][2]==sSite) && !Contains(oSelect.options, eval(valueCol))){
				selected = false;
				oSelect.options[oSelect.options.length] = new Option(eval(textCol), eval(valueCol), false, false);
			}
		if(oSelect.options.length==2)
			oSelect.options[0] = null;
	}
	
	function Contains(options, value){
		for(var count=0;count<options.length;count++)
			if(options[count].value==value)
				return true;
		return false;
	}
	
	function StudyChange(){
		var oStudy = document.forms["selectionheader"]["study"];
		var oSite = document.forms["selectionheader"]["site"];
		if(oSite!=null && oSite.type=="select-one")
			SetContents(oSite,'Sites', 'data[count][2]', 'data[count][2]', oStudy.options[oStudy.selectedIndex].value, '');
		SiteChange();
	}
	
	function SiteChange(){
		var oStudy = document.forms["selectionheader"]["study"];
		var oSite = document.forms["selectionheader"]["site"];
		var oSubject = document.forms["selectionheader"]["subject"];
		if(oSubject!=null && oSubject.type=="select-one"){
			if(oStudy.type=="select-one")
				var lStudyId = oStudy.options[oStudy.selectedIndex].value;
			else
				var lStudyId = oStudy.value;
			if(oSite.type=="select-one")
				var sSite = oSite.options[oSite.selectedIndex].value;
			else
				var sSite = oSite.value;
			SetContents(oSubject, 'Subjects', 'data[count][3]', "''+data[count][1]+'`'+data[count][2]+'`'+data[count][4]", lStudyId, sSite);
		}
	}
	</script>
<%
'	Response.Write "<form id=selectionheader method=get action='http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "'><table bgcolor=#eeeeee cellpadding=3><tr><td>" & HiddenFormElements(Array("study", "site", "subject"))
	Response.Write "<form id=selectionheader method=get action='" & Request.ServerVariables("URL") & "'><table bgcolor=#eeeeee cellpadding=3><tr><td>" & HiddenFormElements(Array("study", "site", "subject"))
	if bMultipleStudies then
		Response.Write "<select name=study onchange='StudyChange();'><option value=''>All Studies</option>" & sStudyHtml & "</select></td><td>"
	else
		Response.Write "<font size=3>" & rsSubjects.Fields("ClinicalTrialName").Value & "</font><input type=hidden name=study value=" & rsSubjects.Fields("ClinicalTrialId").Value & "></td><td>"
	end if
	if bShowSiteList then
		if bMultipleSites then
			Response.Write "<select name=site onchange='SiteChange();'><option value=''>All Sites</option>" & sSiteHtml & "</select></td><td>"
		else
			Response.Write "<font size=3>" & sSite & "</font><input type=hidden name=site value=" & sSite & "></td><td>"
		end if
	end if
	if bShowSubjectList then
		if bMultipleSubjects then
			Response.Write "<select name=subject><option value=''>All Subjects</option>" & sSubjectHtml & "</select></td><td>"
		else
			Response.Write "<font size=3>" & sSubjectLabel & "</font><input type=hidden name=subject value=" & sSubject & "></td><td>"
		end if
	end if
	Response.Write "<input type=submit value='Run Report'>"
	Response.Write "</td></tr></table></form>"
end sub

%>
