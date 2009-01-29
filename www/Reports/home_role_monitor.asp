<v:shapetype id="Bar" fillcolor="blue" strokecolor="blue" strokeweight="0.5pt">
<v:fill type="gradient" />
</v:shapetype>
<!--#include file="report_functions.asp" -->
<%
	WriteUserPanel
	response.write "<br><br>"
	response.write "<table width=""100%"" <font face=""verdana,arial,helvetica"" size=""1"">" 
	response.write "<tr valign=""top""><td width=""33%"">"

	WritePanelStart "Reports"

WriteReportLink  "site_recruitment.asp?PrintDatabase=0" , "Site recruitment"
WriteReportLink  "subject_summary.asp?PrintDatabase=0" , "Subject summary"
WriteReportLink  "schedule_form_summary.asp?PrintDatabase=0" , "eCRF summary"
WriteReportLink  "missing_data_subject.asp?PrintDatabase=0" , "Missing data by subject"
WriteReportLink  "missing_data_eCRF.asp?PrintDatabase=0" , "Missing data by form"
WriteReportLink  "missing_data.asp?PrintDatabase=0" , "Missing data"
WriteReportLink  "laboratory_normal_ranges.asp?PrintDatabase=0" , "Laboratory normal ranges"
WriteReportLink  "Lab_data_abnormal.asp?PrintDatabase=0" , "Out of range lab data"
WriteReportLink  "Discrepancy_Count.asp?PrintDatabase=0" , "Discrepancy Count"
WriteReportLink  "changed_planned_sdv.asp?PrintDatabase=0" , "Done SDVs / changed data"


	  WritePanelEnd

		response.write "<br><br>"
	
		WritePanelStart "References"
%>
<!--#include file="home_references.asp" -->
<%	
		WritePanelEnd


		response.write "<br><br>"
		response.write "<br><br>"
		response.write "<br><br>"
		response.write "<div style=""visibility:hidden"">"
		response.write "<input name=""reporttype""  type=""radio"" checked  ><font style=""font-family:verdana,arial,helvetica;font-size:8pt"">Display / print</font></input><br>"
		response.write "<input name=""reporttype""  type=""radio"" ><font style=""font-family:verdana,arial,helvetica;font-size:8pt"">Excel</font></input><br>"
		response.write "<input name=""reporttype"" type=""radio"" ><font style=""font-family:verdana,arial,helvetica;font-size:8pt"">CSV</font></input><br>"
		response.write "</div>"
	
		WritePanelStart "MACRO reports"
%>
<!--#include file="home_explanation.asp" -->
<%	
		WritePanelEnd	
		
		response.write "</td><td width=""67%"">"
		
Set rsResult = CreateObject("ADODB.Recordset")
Set rsResult1 = CreateObject("ADODB.Recordset")

sQuery = "Select distinct c.clinicaltrialid,clinicaltrialname "
sQuery = sQuery & "from clinicaltrial c  "
sQuery = sQuery & " where c.clinicaltrialid in (" & sPermittedStudyList & ") "
sQuery = sQuery & "order by clinicaltrialname "

rsResult.open sQuery,Connect

do until rsResult.eof 

WriteGroupHeader "Study", rsResult("clinicaltrialname") 

' Display recruitment as chart
sChart = 1
%>
<!--#include file="home_recruitment.asp" -->
<%	

'TA 06/05/2003: i have commented out site specific code to make work temporarily
'sQuery = "Select MIMessageSite,count(mimessageid) as NumberofDisc  "
sQuery = "Select count(mimessageid) as NumberofDisc  "
sQuery = sQuery & "from  MIMessage,clinicaltrial "
sQuery = sQuery & " where clinicaltrialid = '" & rsResult("clinicaltrialid") & "' "
sQuery = sQuery & " and clinicaltrial.clinicaltrialname = MIMessage.MIMessageTrialName " 
'filter on discrepancy current and raised
sQuery = sQuery & " and MIMESSAGETYPE = 0 AND MIMESSAGEHISTORY=0 AND MIMESSAGESTATUS = 1" 
sQuery = sQuery & " and " & replace(sStudySiteSQL, "trialsite.trialsite", "MIMessage.MIMessageSite" )
'sQuery = sQuery & " group by MIMessageSite  "

rsResult1.open sQuery,Connect

if rsResult1.eof then
	 WritePara "There are no discrepancy responses for you to review."
else
		WritePara "There are " & rsResult1("NumberOfDisc") & " discrepancy responses for you to review."
		WritePara "Please select the 'View responded discrepancies' option from the task list on the right."
end if

rsResult1.Close




rsResult.movenext
loop

rsResult.Close





set RsResult = Nothing
set RsResult1 = Nothing

				
		response.write "</td></tr></table>"
%>
