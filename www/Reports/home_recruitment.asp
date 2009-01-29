<%

sQuery = "Select trialsite,count(personid) as Recruitment  "
sQuery = sQuery & "from  trialsubject "
sQuery = sQuery & " where clinicaltrialid = '" & rsResult("clinicaltrialid") & "' "
sQuery = sQuery & " group by trialsite order by recruitment "

rsResult1.open sQuery,Connect

		if rsResult1.eof then
			 WritePara "No subjects recruited yet."
		else
if sChart = 1 then	 'Display chart
	 response.write "<table width=""400px"" height = ""200px"">"
	 nTotal = 0
	 nMax = 0
	 do until rsResult1.eof 
	 		if clng(rsResult1("Recruitment")) > clng(nMax) then
	 			 nMax = clng(rsResult1("Recruitment"))
			end if
			nTotal = nTotal + 1
			rsResult1.movenext
	loop
	response.write "<tr>"
	response.write "<td width= ""80px"">Number of subjects</td>"
	rsResult1.movefirst

	nCount = 0
	do until rsResult1.eof 
		 nCount = nCount + 1
		 response.write "<td height=""90%"" valign=""bottom"" align=""center"">"
		 if nTotal < 20 then
	 	 		nWidth = 20
		 else
		 		 nWidth = nCount/nTotal*400
		 end if
		 response.write rsResult1("Recruitment") & "<br>"
		 response.write "<v:rect type=""#Bar"" style=""width:" & nWidth & "px;height:" & cint(rsResult1("Recruitment")) / cint(nMax) * 150 & "px"" >"
		 response.write "</v:rect>"
		 response.write "</td>"
		 rsResult1.movenext
	loop
	response.write "</tr>"
	response.write "<tr><td></td>"
	rsResult1.movefirst
	do until rsResult1.eof 
		 response.write "<td height=""10%""  align=""center"">"
		 response.write rsResult1("TrialSite")
		 response.write "</td>"
		 rsResult1.movenext
	loop
	response.write "</tr>"
	response.write "</table>"

	
else

		WriteTableStart
		WriteTableRowStart
		WriteHeaderCell "Site"
		WriteHeaderCell "Recruitment"
		WriteTableRowEnd

		do until rsResult1.eof 
			 WriteTableRowStart
			 WriteCell rsResult1("trialsite") 
			 WriteCell rsResult1("recruitment") 
			 WriteTableRowEnd
			 rsResult1.movenext
		loop

		WriteTableEnd

		
end if


end if

rsResult1.Close
%>
