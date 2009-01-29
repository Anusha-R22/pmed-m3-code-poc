<!--#include file=ValidateParameters.asp-->
<%
'validate clinicaltrialid
if (not fnNumeric(request.querystring("clinicaltrialid"))) then
	Response.Write("ERROR:The ClinicalTrialId '" & request.querystring("clinicaltrialid") & "' is not valid")
	Response.End 
end if

response.redirect "d:\my documents\vss\macroreleased\out folder\" & request.querystring("clinicaltrialid") & "\" & request.querystring("clinicaltrialid") & ".1"


%>

