<%

sQuery = "Select *  "
sQuery = sQuery & "from  TrialOffice "

rsResult1.open sQuery,Connect

if not rsResult1.eof then

rsResult1.close

WriteGroupHeader "Data transfer", "" 

sQuery = "Select max(messagetimestamp) as LatestTransferTime  "
sQuery = sQuery & "from  Message "

rsResult1.open sQuery,Connect

if rsResult1.eof then
	 WritePara "Please connect to the server to transfer data."
else
		WritePara "Last connection to the server was on " & cdate(rsResult1("LatestTransferTime")) & "."
'		if datediff("d", rsResult1("LatestTransferTime") , now ) > 2 then
'				WritePara "<b>You have not transferred data for 2 days.  Please transfer data now if possible.</b>"
		if datediff("n", rsResult1("LatestTransferTime") , now ) > 5 then
				WritePara "<b>You have not transferred data for 5 minutes.</b>"
				WritePara "<b>Please transfer data now if possible.</b>"
			end if
end if

end if

rsResult1.Close

%>