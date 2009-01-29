<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'*************************
' Header block
'*************************

sReportTitle = "eForm details"

sIncludeVML = 0 'Don't include VML styles

%>
<!--#include file="report_initialise.asp" -->
<!--#include file="report_open_macro_database.asp" -->
<!--#include file="report_functions.asp" -->
<!--#include file="report_header.asp" -->
<%

'*************************
' Content block
'*************************

nClinicalTrialId = request.querystring("clinicaltrialid")

Set rsResult = CreateObject("ADODB.Recordset")
Set rsResult1 = CreateObject("ADODB.Recordset")
Set rsResult2 = CreateObject("ADODB.Recordset")

sQuery = "Select clinicaltrialname "
sQuery = sQuery & "from clinicaltrial "
sQuery = sQuery & "where clinicaltrial.clinicaltrialid = " & nClinicalTrialId

rsResult.open sQuery,Connect
WriteGroupHeader "Study" , rsResult("clinicaltrialname")
rsResult.close

sQuery = "Select  crfpage.* "
sQuery = sQuery & "from crfpage "
sQuery = sQuery & "where crfpage.clinicaltrialid = " & nClinicalTrialId
sQuery = sQuery & " order by crfpageorder "

rsResult.open sQuery,Connect

do until rsResult.eof 

	 WriteGroupHeader "eForm",  rsResult("crfpagecode")
	 WriteTableStart
	 WriteTableRowStart
	 WriteCell "Code:"
	 WriteCell rsResult("crfpagecode")
	 WriteTableRowEnd 
	 WriteTableRowStart
	 WriteCell "Name:"
	 WriteCell rsResult("crftitle")
	 WriteTableRowEnd 

	 if rsResult("crfpagelabel") > "" then
	 		WriteTableRowStart
	 		WriteCell "Label:"
	 		WriteCell rsResult("crfpagelabel")
	 		WriteTableRowEnd 
	 end if
	 if cint(rsResult("localcrfpagelabel")) > 0 then
	 		WriteTableRowStart
	 		WriteCell "Label:"
	 		WriteCell fLocal (rsResult("localcrfpage"))
	 		WriteTableRowEnd 
	 end if
	 if cint(rsResult("displaynumbers")) > 0 then
	 		WriteTableRowStart
	 		WriteCell ""
	 		WriteCell "Automatic numbering of fields"
	 		WriteTableRowEnd 
	 end if
	 if cint(rsResult("hideifinactive")) > 0 then
	 		WriteTableRowStart
	 		WriteCell ""
	 		WriteCell "Hidden if inactive"
	 		WriteTableRowEnd 
	 end if
	 WriteTableRowStart
	 WriteCell "Page size:"
	 WriteCell fPageSize (rsResult("eformwidth") )
	 WriteTableRowEnd 

	 WriteTableEnd
	 
	 sQuery = "Select visitcode,visitname "
	 sQuery = sQuery & "from studyvisit,studyvisitcrfpage "
	 sQuery = sQuery & "where studyvisit.clinicaltrialid = " & nClinicalTrialId
	 sQuery = sQuery & "  and studyvisitcrfpage.crfpageid = " & rsResult("crfpageid")
	 sQuery = sQuery & " and  studyvisit.clinicaltrialid = studyvisitcrfpage.clinicaltrialid "
	 sQuery = sQuery & " and  studyvisit.visitid = studyvisitcrfpage.visitid "
	 sQuery = sQuery & " order by visitorder "

	 rsResult1.open sQuery,Connect
	 if rsResult1.eof then
	 		WritePara "<B>This eForm is not used in any visits.</b>"
	 else
	 		sVisits =  "<b>Used in visits: "
	 	 do until rsResult1.eof 
			sVisits = sVisits &  rsResult1("visitcode") & ","
				rsResult1.movenext
			loop
			WritePara left(sVisits,len(sVisits) - 1) & "</b>"
	 end if
	 rsResult1.close

	 sQuery = "Select crfelement.*,dataitemcode,dataitemname,datatypename,derivation,datatype,dataitemformat,dataitemlength "
	 sQuery = sQuery & "from crfelement,dataitem,datatype "
	 sQuery = sQuery & "where crfelement.clinicaltrialid = " & nClinicalTrialId
	 sQuery = sQuery & "  and crfelement.crfpageid = " & rsResult("crfpageid")
	 sQuery = sQuery & " and  crfelement.clinicaltrialid = dataitem.clinicaltrialid "
	 sQuery = sQuery & " and  crfelement.dataitemid = dataitem.dataitemid "
	 sQuery = sQuery & " and  dataitem.datatype = datatype.datatypeid "
	 sQuery = sQuery & " order by fieldorder,qgroupfieldorder "

	 rsResult1.open sQuery,Connect
	 if rsResult1.eof then
	 		' Write nothing
	 else
	 WriteTableStart
	 WriteTableRowStart
	 WriteHeaderCell "Field"
	 WriteHeaderCell "Question"
	 WriteHeaderCell ""
	 WriteHeaderCell "Properties"

	 WriteTableRowEnd

	 	 nQGroupId = 0
	 	 do until rsResult1.eof 
		 
		  if cint(nQGroupId) <> cint(rsResult1("OwnerQGroupId")) and cint(rsResult1("OwnerQGroupId")) > cint(0) then
	 			 sQuery = "Select * "
	 			 sQuery = sQuery & "from eformqgroup "
	 			 sQuery = sQuery & "where eformqgroup.clinicaltrialid = " & nClinicalTrialId
	 			 sQuery = sQuery & "  and eformqgroup.crfpageid = " & rsResult("crfpageid")
	 			 sQuery = sQuery & " and  eformqgroup.qgroupid = " & rsResult1("OwnerQGroupId")
	 			 rsResult2.open sQuery,Connect

				 WriteTableRowStart
				 WriteCell ""
				 WriteCell "**"
				 WriteCell "Question group"
				 WriteCell "Display rows: " & rsResult2("displayrows") & " | " & "Initial rows: " & rsResult2("initialrows")  & " | " &  "Min repeats: " & rsResult2("minrepeats")  & " | " & "Max repeats: " & rsResult2("maxrepeats")
				 WriteTableRowEnd		 
				 rsResult2.close
			end if
			nQGroupId = rsResult1("OwnerQGroupId")
		 
		 	sProperties1 = rsResult1("datatypename")	
	 		if rsResult1("dataitemformat") > "" then 
	 			 sProperties1 = sProperties1 & " (" & rsResult1("dataitemlength") & ")"
	 		end if
	 		sProperties1 = sProperties1 & " | "
			if cint(rsResult1("requirecomment")) = 1 then
				 sProperties1 = sProperties1 & "RFC | "
			end if
			if cint(rsResult1("mandatory")) = 1 then
				 sProperties1 = sProperties1 & "Mandatory | "
			end if
			if cint(rsResult1("hidden")) = 1 then
				 sProperties1 = sProperties1 & "Hidden | "
			end if
'		 	if cint(rsResult1("displaylength")) <> cint(rsResult1("dataitemlength")) then
'				 sProperties1 = sProperties1 & "Display length: " & rsResult1("displaylength") & " | "
'			end if
		 	if rsResult1("rolecode") <> "0" and rsResult1("rolecode") > "" then
				 sProperties1 = sProperties1 & "Role: " & rsResult1("rolecode") & " | "
			end if
			if cint(rsResult1("optional")) = 1 then
				 sProperties1 = sProperties1 & "Optional | "
			end if
		 	if cint(rsResult1("localflag")) = 1 then
				 sProperties1 = sProperties1 & "Local | "
			end if

				 sProperties1 = left(sProperties1, len(sProperties1) -2 )
			
			sProperties2 = ""
			sProperties3 = ""
			if rsResult1("skipcondition") > "" then
				 	sProperties2 = sProperties2 & "Only collect if:" & rsResult1("skipcondition")
			end if
			
			if rsResult1("derivation") > "" then
				 if sProperties2 = "" then
				 	sProperties2 = sProperties2 & "Derivation:" & rsResult1("derivation")
				 else
					 	sProperties3 = sProperties3 & "Derivation:" & rsResult1("derivation")			 
				 end if
			end if
			
		 	WriteTableRowStart
			if cint(rsResult1("QGroupFieldOrder")) > 0 then
						WriteCell rsResult1("fieldorder") & "." & rsResult1("QGroupFieldOrder")
			else
						WriteCell rsResult1("fieldorder")
			end if
			WriteCell rsResult1("dataitemcode")
			WriteCell rsResult1("dataitemname")
			WriteCell sProperties1
			WriteTableRowEnd
			if sProperties2 > "" then
				 WriteTableRowStart
				 WriteCell ""
				 WriteCell ""
				 WriteCell ""
				 WriteCell sProperties2
				 WriteTableRowEnd
			end if
			if sProperties3 > "" then
				 WriteTableRowStart
				 WriteCell ""
				 WriteCell ""
				 WriteCell ""
				 WriteCell sProperties3
				 WriteTableRowEnd
			end if
			rsResult1.movenext
			loop
			WriteTableEnd
	 end if
	 rsResult1.close

 
	 rsResult.movenext
loop



rsResult.Close
set RsResult = Nothing
set RsResult1 = Nothing
set RsResult2 = Nothing

'*************************
' Footer block
'*************************

%>
<!--#include file="report_footer.asp" -->
<!--#include file="report_close.asp" -->