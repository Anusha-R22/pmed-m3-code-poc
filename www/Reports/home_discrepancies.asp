<%

'TA 06/05/2003: i have commented out site specific code to make work temporarily
'sQuery = "Select MIMessageSite,count(mimessageid) as NumberofDisc  "
sQuery = "Select count(mimessageid) as NumberofDisc  "
sQuery = sQuery & "from  MIMessage,clinicaltrial "
sQuery = sQuery & " where clinicaltrialid = '" & rsResult("clinicaltrialid") & "' "
sQuery = sQuery & " and clinicaltrial.clinicaltrialname = MIMessage.MIMessageTrialName " 
'filter on discrepancy current and raised
sQuery = sQuery & " and MIMESSAGETYPE = 0 AND MIMESSAGEHISTORY=0 AND MIMESSAGESTATUS = 0" 

sQuery = sQuery & " and " & replace(sStudySiteSQL, "trialsite.trialsite", "MIMessage.MIMessageSite" )
'sQuery = sQuery & " group by MIMessageSite  "
rsResult1.open sQuery,Connect

if rsResult1.eof then
	 WritePara "There are no raised discrepancies."
else
		WritePara "There are " & rsResult1("NumberOfDisc") & " raised discrepancies."
		WritePara "Please select the 'View raised discrepancies' option from the task list on the right."
end if

rsResult1.Close
%>
