<% 
'	 Create objects
	set oUser = server.CreateObject("MACROUSERBS30.MACROUser")
	
'	Initialise user object with serialised string
	if session("ssUser") > "" then				 'WWW
		oUser.setstate(session("UserObject"))
	else																	 'Windows
	  oUser.setstatehex(session("UserObject"))
	end if

' Retrieve info from object
	sUserName = oUser.UserName
	sUserNameFull = oUser.UserNameFull
	sConnectionString = oUser.CurrentDBConString 
	sUserRole = oUser.UserRole
	sLoginDate = oUser.LastLogin
	sDatabase = oUser.DatabaseCode
	sDatabaseType = oUser.Database.DatabaseType
	
' Get list of study ids which user is allowed to access
	sPermittedStudyList = ""
	for each oStudy in oUser.GetAllStudies
	  sPermittedStudyList = sPermittedStudyList & "," & oStudy.StudyId
	next 
	if sPermittedStudyList > "" then
	  sPermittedStudyList = right(sPermittedStudyList, len(sPermittedStudyList) - 1 ) 
	end if

'  Get study and site SQL 
	sStudySiteSQL = oUser.DataLists.StudiesSitesWhereSQL("clinicaltrial.clinicaltrialid", "trialsite.trialsite")
	
' Close object
  set oUser = nothing	 
	
' Set date and time when report was requested
	sReportDate = date() & " " & time()
	
' If report type has been passed, use it
if request.querystring("reporttype") > "" then
	 sReportType = request.querystring("reporttype")
else
		sReportType = 0	' Use HTML as default
end if

' If print data base parameter has been passed, use it
if request.querystring("printdatabase") > "" then
	 sPrintDatabase = request.querystring("printdatabase")
else
		sPrintDatabase = 0	' Default is not to print
end if
%>
