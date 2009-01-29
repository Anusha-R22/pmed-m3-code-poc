Attribute VB_Name = "modTesting"
'------------------------------------------------------------------
' File: modTesting.bas
' Copyright: InferMed Ltd 2004 All Rights Reserved
' Author: Toby Aldridge, September 2004
' Purpose: TESTING PURPOSES ONLY
'------------------------------------------------------------------
' REVISIONS

'------------------------------------------------------------------

Option Explicit

Public Sub TestPassword()

Dim oTest As New ChangeUserDetailsTest
oTest.TestGetUserDetail
oTest.TestUpdateUserPassword
oTest.TestGetUserDetail
End Sub

Sub TestDetails()
Dim oTest As New ChangeUserDetailsTest
oTest.TestGetUserDetail
oTest.TestUpdateUserDetail
oTest.TestGetUserDetail
End Sub



Sub TestGetDetails()
Dim oTest As New ChangeUserDetailsTest
oTest.TestGetUserDetail
End Sub



Sub TestGetSingleDetails()
Dim oTest As New ChangeUserDetailsTest
oTest.TestGetSIngleUserDetail
End Sub
