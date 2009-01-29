<%
	response.write "<br><br>"
	response.write "<table width=""100%"" <font face=""verdana,arial,helvetica"" size=""1"">" 
	response.write "<tr valign=""top""><td width=""33%"">"

	WritePanelStart "Security reports"
WriteReportLink  "functions.asp?PrintDatabase=1" , "Security functions"
WriteReportLink  "rolefunctions.asp?PrintDatabase=1"  , "Roles"
WriteReportLink  "users.asp?PrintDatabase=1"  , "Users"
WriteReportLink  "user_login.asp?PrintDatabase=1"  , "User login activity"
WriteReportLink  "user_login.asp?failed=yes"  , "Failed login attempts"
WriteReportLink  "user_role.asp?PrintDatabase=1"  , "User roles"
WriteReportLink  "password_policy.asp?PrintDatabase=1"  , "Password policy"
	WritePanelEnd
	
	response.write "<br><br>"

	WritePanelStart "Metadata reports"
	
WriteReportLink  "sites.asp?PrintDatabase=1" , "Sites"
WriteReportLink  "study_sites.asp?PrintDatabase=1" , "Study sites"
WriteReportLink  "Units_of_measurement.asp?PrintDatabase=1" , "Units of measurement"
WriteReportLink  "CTC_Schemes.asp?PrintDatabase=1" , "CTC Schemes"
WriteReportLink  "laboratory_normal_ranges.asp?PrintDatabase=1" , "Laboratory normal ranges"
WriteReportLink  "laboratories.asp?PrintDatabase=1" , "Laboratories"
WriteReportLink  "laboratory_sites.asp?PrintDatabase=1" , "Laboratory sites"
WriteReportLink  "test_groups.asp?PrintDatabase=1" , "Clinical test groups"
WriteReportLink  "clinical_tests.asp?PrintDatabase=1" , "Clinical tests"
WriteReportLink  "cdisc_study_list.asp?PrintDatabase=1" , "CDISC"
WriteReportLink  "standard_formats.asp?PrintDatabase=1" , "Standard data formats"
WriteReportLink  "study_phases.asp?PrintDatabase=1" , "Study phases"
WriteReportLink  "validation_types.asp?PrintDatabase=1" , "Validation types"
WriteReportLink  "trial_types.asp?PrintDatabase=1" , "Trial types"
'WriteReportLink  "timezones.asp?PrintDatabase=1" , "Timezones"
WriteReportLink  "countries.asp?PrintDatabase=1" , "Countries"
WriteReportLink  "reserved_words.asp?PrintDatabase=1" , "Reserved words"
WriteReportLink  "study_list.asp?PrintDatabase=1" , "Studies"

	WritePanelEnd
	
	response.write "</td><td width=""33%"">"

	WritePanelStart "Data reports"

WriteReportLink  "data_view_list.asp?PrintDatabase=1" , "Data views"
WriteReportLink  "site_recruitment.asp?PrintDatabase=1" , "Site recruitment"
WriteReportLink  "changed_data.asp?PrintDatabase=1" , "Changed data"
WriteReportLink  "schedule_form_summary.asp?PrintDatabase=1" , "eCRF summary"
WriteReportLink  "subject_summary.asp?PrintDatabase=1" , "Subject summary"
WriteReportLink  "data.asp?PrintDatabase=1" , "Data"
WriteReportLink  "missing_data_site.asp?PrintDatabase=1" , "Missing data by site"
WriteReportLink  "missing_data_subject.asp?PrintDatabase=1" , "Missing data by subject"
WriteReportLink  "missing_data_eCRF.asp?PrintDatabase=1" , "Missing data by form"
WriteReportLink  "missing_data.asp?PrintDatabase=1" , "Missing data"
WriteReportLink  "Lab_data_abnormal.asp?PrintDatabase=1" , "Out of range lab data"

WriteReportLink  "Discrepancy_Count.asp?PrintDatabase=1" , "Discrepancy Count"
WriteReportLink  "changed_planned_sdv.asp?PrintDatabase=1" , "Done SDVs / changed data"



	WritePanelEnd

	response.write "<br><br>"
	response.write "<input name=""reporttype""  type=""radio"" checked  ><font style=""font-family:verdana,arial,helvetica;font-size:8pt"">Display / print</font></input><br>"
	response.write "<input name=""reporttype""  type=""radio"" ><font style=""font-family:verdana,arial,helvetica;font-size:8pt"">Excel</font></input><br>"
	response.write "<input name=""reporttype"" type=""radio"" ><font style=""font-family:verdana,arial,helvetica;font-size:8pt"">CSV</font></input><br>"

		response.write "</td><td width=""33%"">"
	
	WritePanelStart "References"
%>
<!--#include file="home_references.asp" -->
<%	
	WritePanelEnd
	

		
		response.write "</td></tr></table>"
%>
