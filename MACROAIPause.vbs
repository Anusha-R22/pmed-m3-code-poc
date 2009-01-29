'--------------------------------------------------------------------------------
' File:         MACROAIPause.vbs
' Copyright:    InferMed Ltd. 2000. All Rights Reserved
' Author:       Richard Meinesz, June 2002
' Purpose:      Script to update the AutoImportControl table START/STOP field to PAUSE
'---------------------------------------------------------------------------------
'Set the database type and relevant connection strings below
'---------------------------------------------------------------------------------

'******** Set Database Type ********
'For SQL Server use "SQL SERVER"
'For Oracle use "ORACLE"
'For Access use "ACCESS"

sDatabaseType = ""
'______________________________________


'******** Set Wait Interval in Seconds ********
lWaitInterval = 60
'______________________________________


'******** For sql Server ********
sSQLServer = ""
sSQLDatabase = ""
sSQLUserId = ""
sSQLPassword = ""
'______________________________________


'******** For Oracle ********
sOraTNSName = ""
sOraUserId = ""
sOraPassword = ""
'______________________________________


'******** For Access ********
sAccessPath =""
sAccessPassword =""
'______________________________________


Set Connect = CreateObject("ADODB.Connection")

Select Case UCase(sDatabaseType)

Case "SQL SERVER"
Connect.ConnectionString = "PROVIDER=SQLOLEDB;DATA SOURCE=" & sSQLServer & ";DATABASE=" & sSQLDatabase & ";USER ID=" & sSQLUserId & ";PASSWORD=" & sSQLPassword & ";" 

Case "ORACLE"
Connect.ConnectionString = "PROVIDER=MSDAORA;DATA SOURCE=" & sOraTNSName & ";USER ID=" & sOraUserId & ";PASSWORD=" & sOraPassword & ";"

Case "ACCESS"
Connect.ConnectionString = "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & sAccessPath & ";JET OLEDB:DATABASE PASSWORD=" & sAccessPassword & ";"

End Select

Connect.Open

Connect.Execute "UPDATE AutoImportControl SET STARTSTOP='PAUSE', WAITINTERVAL=" & lWaitInterval


