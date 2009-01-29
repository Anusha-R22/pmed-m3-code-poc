<%
'********************************************************************************************************
' Written By:	AN
'
' Revisions
' 06 June 2003 - RS - Add extra ../ to relative documents path
' 10 June 2003 - RS: Get Documents Path from Settings File
'				 RS: Check for file existence, do not display as link if file not found
'********************************************************************************************************
Set rsResult = CreateObject("ADODB.Recordset")
Set rsResult1 = CreateObject("ADODB.Recordset")
Set oFSO = CreateObject("Scripting.FileSystemObject")

' RS 10/06/2003: Create WWW object, get MACRO documents directory
set oIo = server.CreateObject("MACROWWWIO30.clsWWW")
'sDocumentsDirectory = oIo.GetDOcumentsDirectory()
sDocumentsDirectory = "C:\Infermed\DEV\VSS\MACRO 3.0\Documents"

set oIo = Nothing

sQuery = "Select clinicaltrialid,clinicaltrialname "
sQuery = sQuery & "from clinicaltrial "
sQuery = sQuery & "where clinicaltrial.clinicaltrialid > 0 "
sQuery = sQuery & " and clinicaltrial.clinicaltrialid in (" & sPermittedStudyList & ") "
sQuery = sQuery & "order by clinicaltrialname "

rsResult.open sQuery,Connect


do until rsResult.eof 
	 response.write "<tr><td colspan=""2"">" & rsResult("clinicaltrialname") & "</td></tr>"
	 
	 sQuery = "Select * "
	 sQuery = sQuery & "from studydocument "
	 sQuery = sQuery & "where clinicaltrialid = " & rsResult("clinicaltrialid")
	 sQuery = sQuery & "order by documentpath "

	 rsResult1.open sQuery,Connect
	 
	 if rsResult1.eof then
	 		response.write "<tr><td width=""10pt""></td><td>(none)</td></tr>"
	 else
	 do until rsResult1.eof 
	 		sDocumentPath =  server.mappath("../../documents/") & "\" & rsResult1("documentpath")
			'sDocumentPath =  "file:///" & sDocumentsDirectory & "\" & rsResult1("documentpath")
			if oFSO.FileExists(sDocumentPath) then
				' Write Link
				' dph 15/03/2004 - set up link to download asp
		 		'response.write "<tr><td width=""10pt""></td><td><a target=_new href=""" & "file:///" & sDocumentPath & """>" & rsResult1("documentpath") & "</a></td></tr>"	 
		 		response.write "<tr><td width=""10pt""></td><td><a target=_new href=""" & "home_references_download.asp?filename=" & rsResult1("documentpath") & """>" & rsResult1("documentpath") & "</a></td></tr>"	 
			else
				' Write Name Only (no link) as the document is referenced, but was not found where expected
				response.write "<tr><td width=""10pt""></td><td>" & rsResult1("documentpath") & "</td></tr>"	 				
			end if
	 		rsResult1.movenext
	 loop
	 end if	 
	 rsResult1.close

	 rsResult.movenext
loop



rsResult.Close
set RsResult = Nothing
set RsResult1 = Nothing

%>
