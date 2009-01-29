Attribute VB_Name = "modUserFunctions"
'----------------------------------------------------------------------------------------'
'   Copyright:  InferMed Ltd. 1999. All Rights Reserved
'   File:       modUserFunction
'   Author:     Will Casey, Decemeber 1999
'   Purpose:    Assorted User functions used throughout Macro.
'----------------------------------------------------------------------------------------'
'
'----------------------------------------------------------------------------------------'
'   Revisions:
'   NCJ 9 Dec 99 - Added global module access functions
'   NCJ 15 Jan 00 - Added F4006, F4007, F4008
'   NCJ 26 Apr 00 - Added F5013 - F5016
'   NCJ 17 May 00 - Added F5017, F5018
'   NCJ 26 May 00 - Added F3023, F5010
'   NCJ 9 Jun 00  - Added F5019
'   SR3728 WillC 1 Aug 00 -  Added F2009,F2010,F2011,F2012,F2013
'   TA 19/09/2000: New permissions for Lab Test Results
'   NCJ 29 Nov 00 - Added F3024
'   MO 11/6/2002 - Added F1008
'   NCJ 16 Oct 02 - Added F5023
'   ASH 31/10/2002  - Added F3026
'   ic 11/11/2002 - added f5024-f5028
'   ash 18/11/2002 - added gsFnUnRegisterDatabase = "F2014" and 5029 - 5030
'   NCJ 4 Mar 03 - Added F1010 for Batch Validation
'   NCJ 18 Jun 03 - Added F5033 for registration
'   ic 19/07/2005 added clinical coding
'   ic 05/12/2005   added active directory login
'   ic 27/02/2007 issue 2855, enter clinical response is no longer a permission
'----------------------------------------------------------------------------------------'

Public Const gsFnSystemManagement = "F1001"
Public Const gsFnExchange = "F1002"
Public Const gsFnLibraryManagement = "F1003"
Public Const gsFnStudyDefinition = "F1004"
Public Const gsFnDataEntry = "F1005"
Public Const gsFnDataReview = "F1006"
Public Const gsFnCreateDataViews = "F1007"
'Mo 11/6/2002
Public Const gsFnQueryModule = "F1008"
Public Const gsFnBatchDataEntry = "F1009"
' NCJ 4 Mar 03
Public Const gsFnBatchValidation = "F1010"

Public Const gsFnCreateNewUser = "F2001"
Public Const gsFnDisableUser = "F2002"
Public Const gsFnChangeAccessRights = "F2003"
Public Const gsFnMaintRole = "F2004"
Public Const gsFnRegisterDB = "F2005"
Public Const gsFnAssignUserToTrial = "F2006"
Public Const gsFnChangePassword = "F2007"
Public Const gsFnCreateDB = "F2008"
Public Const gsFnChangeSystemProperties = "F2009"
Public Const gsFnViewSystemLog = "F2010"
Public Const gsFnResetPassword = "F2011"
Public Const gsFnViewSiteServerCommunication = "F2012"
Public Const gsFnRestoreDatabase = "F2013"
'ASH 20/11/2002
Public Const gsFnUnRegisterDatabase = "F2014"
'ASH 7/1/2003
Public Const gsFnChangePasswordProperties = "F2015"
'ic 05/12/2005 active directory servers
Public Const gsFnActiveDirectoryServers = "F2016"

Public Const gsFnCreateStudy = "F3001"
Public Const gsFnDelStudy = "F3002"
Public Const gsFnCreateQuestion = "F3003"
Public Const gsFnCopyQuestionFromLib = "F3004"
Public Const gsFnCopyQuestionFromStudy = "F3005"
Public Const gsFnDelQuestion = "F3006"
Public Const gsFnAmendQuestion = "F3007"
Public Const gsFnMaintEForm = "F3008"
Public Const gsFnDelEForm = "F3009"
Public Const gsFnMaintSchedule = "F3010"
Public Const gsFnDelVisit = "F3011"
Public Const gsFnAttachRefDoc = "F3012"
Public Const gsFnRemoveRefDoc = "F3013"
Public Const gsFnAmendArezzo = "F3014"
Public Const gsFnCreateReport = "F3015"
Public Const gsFnDelReport = "F3016"
Public Const gsFnAddEFormToVisit = "F3017"
Public Const gsFnRemoveEFormFromVisit = "F3018"
Public Const gsFnEditStudyDetails = "F3019"
Public Const gsFnCreateEForm = "F3020"
Public Const gsFnCreateVisit = "F3021"
Public Const gsFnMaintVisit = "F3022"
Public Const gsFnGGBArezzoUpdate = "F3023"

Public Const gsFnMaintainRegistration = "F3024"

'For RQGs
Public Const gsFnMaintainQGroups = "F3025"
'ASH 30/10/2002
Public Const gsFnEditQuestionMetadataDescription = "F3026"


Public Const gsFnCreateSite = "F4001"
Public Const gsFnAddSiteToTrialOrTrialToSite = "F4002"
Public Const gsFnRemoveSite = "F4003"
Public Const gsFnDistribNewVersionOfStudyDef = "F4004"
Public Const gsFnChangeTrialStatus = "F4005"
Public Const gsFnImportPatData = "F4006"
Public Const gsFnExportPatData = "F4007"
Public Const gsFnImportStudyDef = "F4008"
'ASH 7/1/2003
Public Const gsFnExportStudyDef = "F4009"
Public Const gsFnLabSiteAdmin = "F4010"
Public Const gsFnExportLab = "F4011"
Public Const gsFnImportLab = "F4012"
Public Const gsFnDistributeLab = "F4013"

Public Const gsFnCreateNewSubject = "F5001"
Public Const gsFnViewData = "F5002"

' NCJ 13 Feb 01, SR 4112 - Changed gsFnFreezeData from 5004 to 5005
Public Const gsFnChangeData = "F5003"
Public Const gsFnLockData = "F5004"
Public Const gsFnFreezeData = "F5005"
Public Const gsFnViewReports = "F5006"
Public Const gsFnMonitorDataReviewData = "F5007"
Public Const gsFnViewCommSettings = "F5008"
Public Const gsFnChangeCommSettings = "F5009"
Public Const gsFnCheckSystemIntegrity = "F5010"

Public Const gsFnCheckAuditIntegrity = "F5012"
Public Const gsFnViewAuditTrail = "F5013"
Public Const gsFnOverruleWarnings = "F5014"
Public Const gsFnAddIComment = "F5015"
Public Const gsFnViewIComments = "F5016"

' NCJ 17/5/00
Public Const gsFnCreateDiscrepancy = "F5017"
Public Const gsFnCreateSDV = "F5018"

' NCJ 9/6/00 SR3458
Public Const gsFnWordTemplates = "F5019"
Public Const gsFnViewSubjectData = "F5020"

'MLM 14/09/01
Public Const gsFnRemoveOwnLocks = "F5021"
Public Const gsFnRemoveAllLocks = "F5022"

' NCJ 16 Oct 02
Public Const gsFnViewSDV = "F5023"

'ic 11/11/2002
Public Const gsFnViewDiscrepancies = "F5024"
Public Const gsFnViewChangesSinceLast = "F5025"
Public Const gsFnChangeDateDisplay = "F5026"
Public Const gsFnSplitScreen = "F5027"
Public Const gsFnViewQuickList = "F5028"
'ASH 20/11/2002
Public Const gsFnUnLockData = "F5029"
Public Const gsFnUnFreezeData = "F5030"

'TA 7/1/2003 - new tranfer data function
Public Const gsFnTransferData = "F5031"

'TA 7/1/2003 - new view lock freeze historyfunction
Public Const gsFnViewLFHistory = "F5032"

Public Const gsFnRegisterSubject = "F5033"

' TA 19/09/2000: New permissions for Lab Test Results
' Permissions are used in LM so new prefix "F6"
Public Const gsFnMaintainLaboratories As String = "F6001"   '"Maintain Laboratories"
Public Const gsFnMaintainCTCSchemes As String = "F6002"     '"Maintain CTC Schemes"
Public Const gsFnMaintainClinicalTests As String = "F6003"  '"Maintain Clinical Tests"
Public Const gsFnMaintainNormalRanges As String = "F6004"   '"Maintain Normal Ranges"
Public Const gsFnMaintainCTC As String = "F6005"            '"Maintain Common Toxicity Criteria"

'ic 27/02/2007 issue 2855, enter clinical response is no longer a permission
'ic 14/07/2005 added clinical coding
'Public Const gsFnEnterClinicalResponse As String = "F6008"
Public Const gsFnCodeClinicalResponse As String = "F6009"
Public Const gsFnChangeClinicalStatus As String = "F6010"
Public Const gsFnValidateClinicalCode As String = "F6011"
Public Const gsFnImportClinicalDictionary As String = "F6012"
Public Const gsFnAutoencodeClinicalResponse As String = "F6013"

