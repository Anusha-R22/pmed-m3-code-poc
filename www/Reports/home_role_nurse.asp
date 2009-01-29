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

' Display recruitment as table
sChart = 0
%>
<!--#include file="home_recruitment.asp" -->
<%	

' Display raised discrepancies
sMessageType = 0
%>
<!--#include file="home_discrepancies.asp" -->
<!--#include file="home_data_transfer_status.asp" -->
<%	

rsResult.movenext
loop

rsResult.Close



set RsResult = Nothing
set RsResult1 = Nothing

				
		response.write "</td></tr></table>"
%>
