<%@ Language=VBScript%>
<%Response.CacheControl = "no-cache"%>
<%Response.AddHeader "Pragma", "no-cache"%>
<%Response.Expires = -1%>
<%
'********************************************************************************************************
' Written By:	AN
'
' Revisions
' 6 June 2003 - RS - Compare resultset values as strings, otherwise type mismatch
' 17/08/2006 - DPH - Age range & effective dates not displaying <= or >= only x to x
'********************************************************************************************************

'*************************
' Header block
'*************************

sReportTitle = "Normal ranges"

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


Set rsResult1 = CreateObject("ADODB.Recordset")
Set rsResult2 = CreateObject("ADODB.Recordset")
Set rsResult3 = CreateObject("ADODB.Recordset")


sQuery = "Select * from Laboratory  "
if request.querystring("laboratorycode") > "" then
sQuery = sQuery & " where laboratorycode = '" & request.querystring("laboratorycode") & "' "
end if
sQuery = sQuery & "order by LaboratoryDescription "


rsResult1.open sQuery,Connect

do until rsResult1.eof 

WriteGroupHeader "Laboratory", rsResult1("LaboratoryDescription") 

sQuery =  "Select n.*, cg.ClinicalTestGroupDescription, ct.ClinicalTestDescription, ct.Unit from normalrange n, ClinicalTest ct, ClinicalTestGroup cg  "
sQuery = sQuery & "where LaboratoryCode ='" & rsResult1("LaboratoryCode") & "' "
sQuery = sQuery & " and n.clinicaltestcode = ct.clinicaltestcode "
sQuery = sQuery & "  and ct.clinicaltestgroupcode = cg.clinicaltestgroupcode "
sQuery = sQuery & "order by ClinicalTestGroupDescription,ClinicalTestDescription "

rsResult2.open sQuery,Connect


WriteTableStart
WriteTableRowStart
WriteHeaderCell  "Test Group"
WriteHeaderCell  "Clinical Test"
WriteHeaderCell  "Gender"
WriteHeaderCell  "Age range"
WriteHeaderCell  "Normal range"
WriteHeaderCell  "Feasible range"
WriteHeaderCell  "Absolute range"
WriteHeaderCell  "Effective dates"
WriteTableRowEnd

do until rsResult2.eof 
WriteTableRowStart
WriteCell rsResult2("ClinicalTestGroupDescription") 
WriteCell rsResult2("ClinicalTestDescription") 
 
select case rsResult2("NormalRangeGender")
case "2"
	 WriteCell "Male"
case "1"
	 WriteCell  "Female"
case else
		WriteCell ""
end select


sRange = ""
' age range
' if full range
if rsResult2("NormalRangeAgeMin") > "" and rsResult2("NormalRangeAgeMax") > "" then
	 sRange = rsResult2("NormalRangeAgeMin") & " to " & rsResult2("NormalRangeAgeMax")
	 ' elseif just Age Max
elseif ((rsResult2("NormalRangeAgeMin") = "") or isnull(rsResult2("NormalRangeAgeMin"))) and rsResult2("NormalRangeAgeMax") > "" then
	 sRange = "<= " & rsResult2("NormalRangeAgeMax")
	 ' elseif just Age Min
elseif rsResult2("NormalRangeAgeMin") > "" and ((rsResult2("NormalRangeAgeMax") = "") or isnull(rsResult2("NormalRangeAgeMax"))) then
	 sRange = ">= " & rsResult2("NormalRangeAgeMin") 
end if

WriteCell  sRange 

sRange = ""
if rsResult2("NormalRangeNormalMin") > "" and rsResult2("NormalRangeNormalMax") > "" then
	 sRange = rsResult2("NormalRangeNormalMin") & " to " & rsResult2("NormalRangeNormalMax")
elseif  rsResult2("NormalRangeNormalMax") > "" then
	 sRange = "<= " & rsResult2("NormalRangeNormalMax")
elseif rsResult2("NormalRangeNormalMin")  > "" then
	 sRange = ">= " & rsResult2("NormalRangeNormalMin") 
end if

WriteCell  sRange 

sRange = ""
if rsResult2("NormalRangePercent") = "1" or rsResult2("NormalRangePercent") = "3" then
	 sPercent = "%"
else
		sPercent = ""
end if
if rsResult2("NormalRangeFeasibleMin") > "" and rsResult2("NormalRangeFeasibleMax") > "" then
	 sRange = rsResult2("NormalRangeFeasibleMin") & sPercent & " to " & rsResult2("NormalRangeFeasibleMax") & sPercent
elseif  rsResult2("NormalRangeFeasibleMax") > "" then
	 sRange = "<= " & rsResult2("NormalRangeFeasibleMax") & sPercent
elseif rsResult2("NormalRangeFeasibleMin")  > "" then
	 sRange = ">= " & rsResult2("NormalRangeFeasibleMin")  & sPercent
end if

WriteCell  sRange 

sRange = ""
if rsResult2("NormalRangePercent") >= "2" then
	 sPercent = "%"
else
		sPercent = ""
end if
if rsResult2("NormalRangeAbsoluteMin") > "" and rsResult2("NormalRangeAbsoluteMax") > "" then
	 sRange = rsResult2("NormalRangeAbsoluteMin") & sPercent & " to " & rsResult2("NormalRangeAbsoluteMax") & sPercent
elseif  rsResult2("NormalRangeAbsoluteMax") > "" then
	 sRange = "<= " & rsResult2("NormalRangeAbsoluteMax") & sPercent
elseif rsResult2("NormalRangeAbsoluteMin") = ""  > "" then
	 sRange = ">= " & rsResult2("NormalRangeAbsoluteMin")  & sPercent
end if

WriteCell  sRange 

sRange = ""
' DPH 22/07/2004 - Show dates (if applicable) for normal range start / end
if rsResult2("NormalRangeEffectiveStart") > "0" and rsResult2("NormalRangeEffectiveEnd") > "0" then
	 sRange = fConvertDate(rsResult2("NormalRangeEffectiveStart")) & " to " & fConvertDate(rsResult2("NormalRangeEffectiveEnd"))
elseif rsResult2("NormalRangeEffectiveStart") = "0" and rsResult2("NormalRangeEffectiveEnd") > "0" then
	 sRange = "<= " & fConvertDate(rsResult2("NormalRangeEffectiveEnd"))
elseif rsResult2("NormalRangeEffectiveStart") > "0" and rsResult2("NormalRangeEffectiveEnd") = "0" then
	 sRange = ">= " & fConvertDate(rsResult2("NormalRangeEffectiveStart"))
end if

WriteCell  sRange 
WriteTableRowEnd
rsResult2.movenext
loop

rsResult2.Close
WriteTableEnd



rsResult1.movenext
loop



rsResult1.Close
set RsResult1 = Nothing
set RsResult2 = Nothing

'*************************
' Footer block
'*************************

%>
  <!--#include file="report_footer.asp" -->
<!--#include file="report_close.asp" -->