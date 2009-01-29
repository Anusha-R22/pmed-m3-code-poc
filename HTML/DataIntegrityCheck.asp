<!--#include file=LocalSettings.txt-->
<!--#include file=OpenDataConnection.txt-->
<!--#include file=StringUtilities.txt-->
<!--#include file=ValidateParameters.asp-->

<%
'-----------------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 2002 All Rights Reserved
'   File:       DataIntegrityCheck.asp
'   Author:     David Hook
'   Purpose:    Used by TrialOffice for the purpose of creating a data integrity
'				check for inspection on the server
'-----------------------------------------------------------------------------------------------'
'
'-----------------------------------------------------------------------------------------------'
'   Revisions:
' DPH 25/04/2002 - Changed UserID to Username in the SQL
' DPH 01/05/2002 - Changed error handling to <> 0
' ic 24/05/2004 added variable checking
'-----------------------------------------------------------------------------------------------'

dim sSQL, nNextLogNo, sTimeNow, bMIOK
dim vCountMessageIdZero, vCountMessageIdOne, vSource
dim rsMIData

on error resume next

if err.number <> 0 then
	Response.Write "ERROR1:" & err.number
	Response.End 
else

	'validate site
	if (not fnValidateSite(request.querystring("SITENAME"))) then
		Response.Write("ERROR:The site '" & request.querystring("SITENAME") & "' does not exist")
		Response.End 
	end if
	'validate USERID
	if not fnValidateUsername(Request.querystring("USERID")) then
		Response.Write("ERROR:The userid '" & Request.querystring("USERID") & "' is not valid")
		Response.End
	end if
	

	Select Case request.querystring("DIVTYPE")
		Case "MIMessage"
			' Deal with the MIMessage Data Integrity Check
			' If have all data
			If Request.querystring("DIVIDZERO") <> "" and Request.querystring("DIVIDONE") <> "" and Request.querystring("COUNTIDZERO") <> "" and Request.querystring("COUNTIDONE") <> "" _
				and Request.querystring("SITENAME") <> "" and Request.querystring("USERID") <> "" Then

				'validate DIVIDZERO
				if (not fnNumeric(request.querystring("DIVIDZERO"))) then
					Response.Write("ERROR:The DIVIDZERO '" & request.querystring("DIVIDZERO") & "' is not valid")
					Response.End 
				end if
				'validate DIVIDONE
				if (not fnNumeric(request.querystring("DIVIDONE"))) then
					Response.Write("ERROR:The DIVIDONE '" & request.querystring("DIVIDONE") & "' is not valid")
					Response.End 
				end if
				'validate COUNTIDZERO
				if (not fnNumeric(request.querystring("COUNTIDZERO"))) then
					Response.Write("ERROR:The COUNTIDZERO '" & request.querystring("COUNTIDZERO") & "' is not valid")
					Response.End 
				end if
				'validate COUNTIDONE
				if (not fnNumeric(request.querystring("COUNTIDONE"))) then
					Response.Write("ERROR:The COUNTIDONE '" & request.querystring("COUNTIDONE") & "' is not valid")
					Response.End 
				end if

				' Get DI data from server
				sSQL = "SELECT Count(MIMessageId) AS CountID, MIMessageSource FROM MIMessage WHERE MIMessageSite = '" & Request.querystring("SITENAME") & "' GROUP BY MIMessageSource"
				Set rsMIData = MACROCnn.Execute(sSQL) 
				if err.number <> 0 then
					Response.Write "ERROR2:" & err.number
					Response.End
				end if
				
				' Initialise variables
				vCountMessageIdZero = 0
				vCountMessageIdOne = 0

				Do While Not rsMIData.EOF

					vSource = rsMIData("MIMessageSource")
					Select Case cint(vSource)
					    Case 0
					        vCountMessageIdZero = rsMIData(0).Value
					    Case 1
					        vCountMessageIdOne = rsMIData(0).Value
					End Select

					rsMIData.MoveNext
					if err.number <> 0 then
						Response.Write "ERROR3:" & err.number  
						Response.End
					end if
				Loop
				rsMIData.Close
				Set rsMIData = Nothing
				
				' test if checks succeed
				If cdbl(vCountMessageIdZero) = cdbl(Request.querystring("COUNTIDZERO")) And cdbl(vCountMessageIdOne) = cdbl(Request.querystring("COUNTIDONE")) Then
					bMIOK = true
				else
					bMIOK = false
				end if
								
				' Create INSERT SQL Statement for Log table
				sTimeNow = ConvertLocalNumToStandard(Cstr(CDbl(Now)))
				nNextLogNo = NextLogNumber(sTimeNow)
				
				sSQL = "INSERT INTO LogDetails (LogDateTime,LogNumber,TaskId,LogMessage,UserName,LogDateTime_TZ,Location,Status) VALUES (" & sTimeNow & "," & nNextLogNo & ",'DataIntegrity',"
				if bMIOK = true then
					sSQL = sSQL & "'DI MIMessage. Site " & Request.querystring("SITENAME") & " has passed MIMessage data integrity check. "
				else
					sSQL = sSQL & "'DI MIMessage Error. Site " & Request.querystring("SITENAME") & " has failed MIMessage data integrity check. "
				end if
				sSQL = sSQL & "Site DI Counts are Source 0 = " & Request.querystring("COUNTIDZERO") & " Source 1 = " & Request.querystring("COUNTIDONE")
				sSQL = sSQL & " Site checksums are Source 0 ID = " & Request.querystring("DIVIDZERO") & " and Source1 ID = " & Request.querystring("DIVIDONE")
				sSQL = sSQL & "','" & Request.querystring("USERID") & "'," & 0 & ",'Local'," & 0 & ")"
				
				MACROCnn.Execute sSQL
				if err.number <> 0 then
					Response.Write "ERROR4:" & err.number
					Response.End
				end if

			Else
				Response.Write "ERROR5:Some input parameters missing"
				Response.End
			End If
						
		Case "History"
			' Write The Subject Data Integrity Value to the Log Table
			' If have all data
			If Request.querystring("DIVSUBJECT") <> "" and Request.querystring("MAXTIMESTAMP") <> "" and Request.querystring("SUBJECTCOUNT") <> "" and Request.querystring("SITENAME") <> "" and Request.querystring("USERID") <> "" Then
				'validate DIVSUBJECT
				if (not fnNumeric(request.querystring("DIVSUBJECT"))) then
					Response.Write("ERROR:The DIVSUBJECT '" & request.querystring("DIVSUBJECT") & "' is not valid")
					Response.End 
				end if
				'validate MAXTIMESTAMP
				if (not fnNumeric(request.querystring("MAXTIMESTAMP"))) then
					Response.Write("ERROR:The MAXTIMESTAMP '" & request.querystring("MAXTIMESTAMP") & "' is not valid")
					Response.End 
				end if
				'validate SUBJECTCOUNT
				if (not fnNumeric(request.querystring("SUBJECTCOUNT"))) then
					Response.Write("ERROR:The SUBJECTCOUNT '" & request.querystring("SUBJECTCOUNT") & "' is not valid")
					Response.End 
				end if
			
			
				' Create INSERT SQL Statement for Log table
				sTimeNow = ConvertLocalNumToStandard(Cstr(CDbl(Now)))
				nNextLogNo = NextLogNumber(sTimeNow)
				
				sSQL = "INSERT INTO LogDetails (LogDateTime,LogNumber,TaskId,LogMessage,UserName,LogDateTime_TZ,Location,Status) VALUES (" & sTimeNow & "," & nNextLogNo & ",'DataIntegrity',"
				sSQL = sSQL & "'DI Subject data. Site " & Request.querystring("SITENAME") & " has generated a DI check count of " & Request.querystring("SUBJECTCOUNT") & ". Timestamp = " & ConvertLocalNumToStandard(Request.querystring("MAXTIMESTAMP")) & " Sum = " & Request.querystring("DIVSUBJECT") & "','" & Request.querystring("USERID") & "'," & 0 & ",'Local'," & 0 & ")"
				'Response.Write sSQL ' ***DPH		
				MACROCnn.Execute sSQL
				if err.number <> 0 then
					Response.Write "ERROR6:" & err.number
					Response.End
				end if
			Else
				Response.Write "ERROR7:Some input parameters missing"
				Response.End
			End If					

		Case "Connect"
			Response.Write "SUCCESS"
			Response.End 
			
		Case Else
			' unknown 
			Response.Write "ERROR8:Unknown DIVTYPE"
			Response.End

	End Select

	if err.number <> 0 then
		response.write "ERROR9:" & err.number
	else
		' ASP page has run successfully
		response.write "SUCCESS"
	end if
end if

function NextLogNumber(sTimeNow)

	dim rsLogDetails
	
	' Get log number of records
	sSQL = "SELECT Count(LogNumber) From LogDetails WHERE LogDateTime = " & sTimeNow

	'assess the number of records and set the LogNumber for this entry (nLogNumber)
	Set rsLogDetails = MACROCnn.Execute(sSQL)
	NextLogNumber = rsLogDetails(0).Value
	rsLogDetails.Close

	set rsLogDetails = nothing
end function
%>

<!--#include file=CloseDataConnection.txt-->
